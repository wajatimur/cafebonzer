VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CafeBonzer - Setting"
   ClientHeight    =   7365
   ClientLeft      =   1755
   ClientTop       =   2085
   ClientWidth     =   8340
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
   Icon            =   "FrmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin CafeBonzer.XpButton SetBtn 
      Height          =   435
      Index           =   1
      Left            =   7710
      TabIndex        =   63
      ToolTipText     =   "Save settings and exit."
      Top             =   6840
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
      MICON           =   "FrmSet.frx":08CA
      PICN            =   "FrmSet.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6600
      Left            =   75
      TabIndex        =   14
      Top             =   45
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   11642
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   443
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmSet.frx":0E80
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "IDFrame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "NetFrame"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "InfoFrame"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MpassFrame"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Financial"
      TabPicture(1)   =   "FrmSet.frx":0E9C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HargaFrame"
      Tab(1).Control(1)=   "OverFrame"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Employee"
      TabPicture(2)   =   "FrmSet.frx":0EB8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "UaHrd1"
      Tab(2).Control(1)=   "EmpUaBtn"
      Tab(2).Control(2)=   "EmpBtn(1)"
      Tab(2).Control(3)=   "EmpBtn(0)"
      Tab(2).Control(4)=   "Lv1"
      Tab(2).Control(5)=   "Opt1(2)"
      Tab(2).Control(6)=   "Opt1(1)"
      Tab(2).Control(7)=   "Opt1(0)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Security"
      TabPicture(3)   =   "FrmSet.frx":0ED4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Appearance"
      TabPicture(4)   =   "FrmSet.frx":0EF0
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.CheckBox Opt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow access to Settings."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   -74565
         TabIndex        =   67
         Top             =   3735
         Width           =   2625
      End
      Begin VB.CheckBox Opt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow access to Statistic."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   -74565
         TabIndex        =   66
         Top             =   4035
         Width           =   2730
      End
      Begin VB.CheckBox Opt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow unlock Client."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   -74565
         TabIndex        =   65
         Top             =   4365
         Width           =   2730
      End
      Begin VB.Frame HargaFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pricing :"
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
         Height          =   6060
         Left            =   -71055
         TabIndex        =   42
         Top             =   390
         Width           =   4125
         Begin CafeBonzer.XpButton PriBtn 
            Height          =   360
            Index           =   0
            Left            =   3615
            TabIndex        =   69
            Top             =   5625
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
            MICON           =   "FrmSet.frx":0F0C
            PICN            =   "FrmSet.frx":0F28
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox PriPmTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   855
            MaxLength       =   5
            TabIndex        =   48
            ToolTipText     =   "Example : 0.03 for 3 cent per minute."
            Top             =   645
            Width           =   570
         End
         Begin VB.TextBox PriPmTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2580
            MaxLength       =   5
            TabIndex        =   47
            ToolTipText     =   "Initial price."
            Top             =   645
            Width           =   735
         End
         Begin VB.TextBox PriTxt1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2475
            TabIndex        =   46
            Top             =   4050
            Width           =   1230
         End
         Begin VB.TextBox PriTxt1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2475
            TabIndex        =   45
            Top             =   3600
            Width           =   1230
         End
         Begin VB.CheckBox PriChk1 
            Appearance      =   0  'Flat
            Caption         =   "Round Up Price"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   135
            TabIndex        =   43
            Top             =   5640
            Width           =   2220
         End
         Begin CafeBonzer.Line3D PriuLine1 
            Height          =   45
            Left            =   60
            TabIndex        =   44
            Top             =   5505
            Width           =   3990
            _ExtentX        =   7038
            _ExtentY        =   79
            horizon         =   -1  'True
         End
         Begin MSComctlLib.ListView PriLV1 
            Height          =   1755
            Left            =   255
            TabIndex        =   49
            Top             =   1740
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   3096
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Skema harga"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Harga\Minit"
               Object.Width           =   2469
            EndProperty
         End
         Begin CafeBonzer.XpButton PriBtn 
            Height          =   360
            Index           =   1
            Left            =   3180
            TabIndex        =   70
            Top             =   5625
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
            MICON           =   "FrmSet.frx":14C2
            PICN            =   "FrmSet.frx":14DE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RM :"
            Height          =   195
            Left            =   390
            TabIndex        =   55
            Top             =   690
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/Minute  +"
            Height          =   195
            Left            =   1500
            TabIndex        =   54
            Top             =   690
            Width           =   885
         End
         Begin VB.Label PriLBL2 
            BackStyle       =   0  'Transparent
            Caption         =   "Price per Minute :"
            Height          =   255
            Left            =   780
            TabIndex        =   53
            Top             =   4080
            Width           =   1530
         End
         Begin VB.Label PriLBL1 
            BackStyle       =   0  'Transparent
            Caption         =   "Scheme Name :"
            Height          =   255
            Left            =   780
            TabIndex        =   52
            Top             =   3630
            Width           =   1395
         End
         Begin VB.Label PriHdr1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Normal Pricing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   150
            TabIndex        =   51
            Top             =   255
            Width           =   3825
         End
         Begin VB.Label PriHrd2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Additional Pricing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   150
            TabIndex        =   50
            Top             =   1335
            Width           =   3825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Global Employee Setting :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2160
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   7980
         Begin VB.CheckBox GwsOpt1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Log workers activities."
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   41
            Top             =   315
            Width           =   2625
         End
         Begin VB.CheckBox GwsOpt1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Allow changes price."
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   165
            TabIndex        =   40
            Top             =   615
            Width           =   2625
         End
         Begin VB.CheckBox GwsOpt1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Allow cancel after 10 sec."
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   165
            TabIndex        =   39
            Top             =   945
            Width           =   2625
         End
      End
      Begin VB.Frame MpassFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Master Password :"
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
         Height          =   1785
         Left            =   165
         TabIndex        =   30
         Top             =   4605
         Width           =   3720
         Begin VB.TextBox MpassTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1545
            MaxLength       =   20
            PasswordChar    =   "l"
            TabIndex        =   59
            Top             =   1245
            Width           =   1965
         End
         Begin VB.TextBox MpassTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1545
            MaxLength       =   20
            PasswordChar    =   "l"
            TabIndex        =   4
            Top             =   765
            Width           =   1965
         End
         Begin VB.TextBox MpassTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1545
            MaxLength       =   20
            TabIndex        =   3
            Top             =   285
            Width           =   1965
         End
         Begin VB.Label MpassLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retype :"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   58
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label MpassLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   32
            Top             =   795
            Width           =   945
         End
         Begin VB.Label MpassLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username :"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   31
            Top             =   330
            Width           =   1005
         End
      End
      Begin VB.Frame OverFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Overhead && Account Setting :"
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
         Height          =   6045
         Left            =   -74835
         TabIndex        =   26
         Top             =   390
         Width           =   3660
         Begin VB.ComboBox OverCmb1 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            ItemData        =   "FrmSet.frx":1878
            Left            =   2430
            List            =   "FrmSet.frx":1882
            TabIndex        =   11
            Text            =   "AM"
            Top             =   2925
            Width           =   600
         End
         Begin VB.TextBox OverSesTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   1
            Left            =   1845
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "30"
            Top             =   2925
            Width           =   450
         End
         Begin VB.TextBox OverSesTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   0
            Left            =   1155
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "12"
            Top             =   2925
            Width           =   450
         End
         Begin VB.TextBox OvhTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   0
            Left            =   2100
            TabIndex        =   6
            Text            =   "800"
            Top             =   675
            Width           =   930
         End
         Begin VB.TextBox OvhTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   1
            Left            =   2100
            TabIndex        =   7
            Text            =   "350"
            Top             =   1065
            Width           =   930
         End
         Begin VB.TextBox OvhTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   2
            Left            =   2100
            TabIndex        =   8
            Text            =   "20"
            Top             =   1455
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   195
            Left            =   1680
            TabIndex        =   36
            Top             =   2955
            Width           =   75
         End
         Begin VB.Label OverLb4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Auto Session Close On :"
            Height          =   195
            Left            =   285
            TabIndex        =   35
            Top             =   2580
            Width           =   2085
         End
         Begin VB.Label OverHdr2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Account Settings"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   150
            TabIndex        =   34
            Top             =   2115
            Width           =   3345
         End
         Begin VB.Label OverHdr1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Monthly Overhead"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   150
            TabIndex        =   33
            Top             =   285
            Width           =   3375
         End
         Begin VB.Label OverLb1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Premise Rental :"
            Height          =   195
            Left            =   255
            TabIndex        =   29
            Top             =   705
            Width           =   1425
         End
         Begin VB.Label OverLb2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Electric/Water Bills :"
            Height          =   195
            Left            =   255
            TabIndex        =   28
            Top             =   1095
            Width           =   1740
         End
         Begin VB.Label OverLb3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Others :"
            Height          =   195
            Left            =   270
            TabIndex        =   27
            Top             =   1470
            Width           =   705
         End
      End
      Begin VB.Frame InfoFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Information :"
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
         Height          =   2535
         Left            =   165
         TabIndex        =   22
         Top             =   1980
         Width           =   3735
         Begin VB.TextBox InfoTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   2
            Left            =   675
            TabIndex        =   2
            Text            =   "keep our pc clean"
            Top             =   1965
            Width           =   2865
         End
         Begin VB.TextBox InfoTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   0
            Left            =   690
            TabIndex        =   0
            Text            =   "good cybercafe"
            Top             =   555
            Width           =   2865
         End
         Begin VB.TextBox InfoTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Index           =   1
            Left            =   690
            TabIndex        =   1
            Text            =   "owner@cybercafe.com"
            Top             =   1260
            Width           =   2865
         End
         Begin VB.Label InfoLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motto :"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   25
            Top             =   1695
            Width           =   600
         End
         Begin VB.Label InfoLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cybercafes Name :"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   285
            Width           =   1665
         End
         Begin VB.Label InfoLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail :"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   23
            Top             =   1005
            Width           =   675
         End
      End
      Begin VB.Frame NetFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Networking :"
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
         Height          =   1860
         Left            =   4020
         TabIndex        =   20
         Top             =   390
         Width           =   4035
         Begin VB.TextBox NetPassTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   9
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2055
            PasswordChar    =   "l"
            TabIndex        =   62
            ToolTipText     =   "Communication port between the server and the client, 56266 is the default value."
            Top             =   1305
            Width           =   1830
         End
         Begin VB.TextBox NetPassTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   9
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   2055
            PasswordChar    =   "l"
            TabIndex        =   60
            ToolTipText     =   "Default password for clients, if this field left empty, it will be set same as master password."
            Top             =   780
            Width           =   1830
         End
         Begin VB.TextBox NetPortTxt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   2295
            MaxLength       =   5
            TabIndex        =   5
            Text            =   "56266"
            ToolTipText     =   "Communication port between the server and the client, 56266 is the default value."
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label NetLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retype Password :"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   61
            Top             =   1350
            Width           =   1605
         End
         Begin VB.Label NetLbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Client Password :"
            Height          =   420
            Index           =   1
            Left            =   210
            TabIndex        =   37
            Top             =   750
            Width           =   1380
         End
         Begin VB.Label NetLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Port :"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   21
            Top             =   315
            Width           =   975
         End
      End
      Begin VB.Frame IDFrame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registered To :"
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
         Height          =   1530
         Left            =   180
         TabIndex        =   15
         Top             =   390
         Width           =   3720
         Begin CafeBonzer.XpButton RegBtn 
            Height          =   390
            Left            =   2475
            TabIndex        =   57
            Top             =   1035
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   688
            TX              =   "Verify"
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
            MICON           =   "FrmSet.frx":188E
            PICN            =   "FrmSet.frx":18AA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtNombor 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1335
            TabIndex        =   17
            Top             =   645
            Width           =   2220
         End
         Begin VB.TextBox TxtNama 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1335
            TabIndex        =   16
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label RegLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liscence  :"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   19
            Top             =   690
            Width           =   915
         End
         Begin VB.Label RegLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   285
            Width           =   630
         End
      End
      Begin MSComctlLib.ListView Lv1 
         Height          =   2715
         Left            =   -74865
         TabIndex        =   56
         Top             =   465
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4789
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "Iml"
         SmallIcons      =   "Iml"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salary"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Password"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Access"
            Object.Width           =   2540
         EndProperty
      End
      Begin CafeBonzer.XpButton EmpBtn 
         Height          =   360
         Index           =   0
         Left            =   -74430
         TabIndex        =   71
         ToolTipText     =   "Add new employee."
         Top             =   6105
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
         MICON           =   "FrmSet.frx":1E44
         PICN            =   "FrmSet.frx":1E60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton EmpBtn 
         Height          =   360
         Index           =   1
         Left            =   -74865
         TabIndex        =   72
         ToolTipText     =   "Delete selected employee."
         Top             =   6105
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
         MICON           =   "FrmSet.frx":23FA
         PICN            =   "FrmSet.frx":2416
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton EmpUaBtn 
         Height          =   360
         Left            =   -67410
         TabIndex        =   73
         ToolTipText     =   "Click to save user access settings."
         Top             =   3675
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
         MICON           =   "FrmSet.frx":27B0
         PICN            =   "FrmSet.frx":27CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label UaHrd1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " User setting"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74850
         TabIndex        =   68
         Top             =   3345
         Width           =   7920
      End
   End
   Begin CafeBonzer.Label3D Label3D2 
      Height          =   225
      Left            =   705
      TabIndex        =   13
      Top             =   7005
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   4210752
      Caption         =   "Copyright 1996-2003 Nematix Technology"
      BackColor       =   12632256
   End
   Begin CafeBonzer.Label3D Label3D1 
      Height          =   210
      Left            =   690
      TabIndex        =   12
      Top             =   6735
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   4210752
      Caption         =   "CafeBonzer v1.7"
      BackColor       =   12632256
   End
   Begin CafeBonzer.XpButton SetBtn 
      Height          =   435
      Index           =   0
      Left            =   7155
      TabIndex        =   64
      ToolTipText     =   "Cancel all settings and exit."
      Top             =   6840
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
      MICON           =   "FrmSet.frx":2D66
      PICN            =   "FrmSet.frx":2D82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "FrmSet.frx":331C
      Top             =   6750
      Width           =   480
   End
End
Attribute VB_Name = "FrmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DemoMode As Boolean

Function InitReg() As Boolean
    'Menyediakan algoritma bagi pengiraan kod daftar
    a1 = Len(TxtNama)      'ambil panjang nama
    a2 = a1 * 5            'panjang nama darab dengan 5
    a3 = a2 * a1           'hasil darab, di darabkan balik dengan panjang nama
    a4 = Left(TxtNama, 1)  'ambil huruf paling hujung sebelah kiri
    a5 = Right(TxtNama, 1) 'ambil huruf paling hujung sebelah kanan
    a6 = "0v10"            'untuk versi = versi 0.10
    
    'menambahkan dan mengaturkan algoritma kod daftar
    genstr = a1 & a2 & a3 & a4 & a5 & a6
    idstr = TxtNombor
    
    'membandingkan kod daftar dengan nama
    If LCase(Trim(genstr)) = LCase(Trim(idstr)) Then
    InitReg = True
    Else:
    InitReg = False
    End If
End Function

Private Sub EmpBtn_Click(Index As Integer)
    Select Case ButtonIndex
    Case 0
        FrmAddWorker.Show vbModal
    Case 1
        If Lv1.ListItems.Count = 0 Then Exit Sub
        If Lv1.SelectedItem.Text = "" Then Exit Sub
        nm = Lv1.SelectedItem.Tag
        lret = MsgBox("Delete " & Lv1.SelectedItem.Text & " ?", vbOKCancel, CbMsgWarn)
        If lret = vbCancel Then Exit Sub
        uSDBe.DataRemove "pekerja-list", "nama", nm
        Lv2.ListItems.Remove (Lv1.SelectedItem.Index)
        Lv1.ListItems.Remove (Lv1.SelectedItem.Index)
    End Select
End Sub

Private Sub EmpUaBtn_Click()
    Dim TmpAk As String
    If Lv1.ListItems.Count = 0 Then Exit Sub
    If Lv1.SelectedItem.Text = "" Then Exit Sub
    
    For g = 1 To 3
        If Opt1(g - 1).Value = 1 Then
            TmpAk = TmpAk & "1"
        Else
            TmpAk = TmpAk & "0"
        End If
    Next g
    Lv1.SelectedItem.SubItems(4) = TmpAk
    uSDBe.DataEdit "pekerja-list", "akses", "nick", Lv1.SelectedItem.SubItems(2), TmpAk, True, True
End Sub

Private Sub Form_Load()
On Error GoTo ErrInt
    Dim lItm As ListItem, acsTime As String, strNm As String
    
    'SetAmbil "logaktiviti", 1 '= "" Then SetSimpan "logaktiviti", 1
    'SetAmbil "tukarharga", 0 '= "" Then SetSimpan "tukarharga", 0
    'SetAmbil ("roundup"), 0  '= "" Then SetSimpan "roundup", 0
    NumOnly NetPortTxt
    NumOnly OvhTxt(0)
    NumOnly OvhTxt(1)
    NumOnly OvhTxt(2)
    NumOnly OverSesTxt(0)
    NumOnly OverSesTxt(1)
    
    TxtNama = SetAmbil("namadaftar")                    'ambil data nama bagi pengguna
    TxtNombor = SetAmbil("nombordaftar")
    
    MpassTxt(0) = SetAmbil("mu", "admin")
    MpassTxt(1) = SetAmbil("mp")
    MpassTxt(2) = MpassTxt(1)
    
    InfoTxt(0) = SetAmbil("namacc")                     'nama kedai cc
    InfoTxt(1) = SetAmbil("emailpengguna")              'email pengguna
    InfoTxt(2) = SetAmbil("tajukatas")                       'tajuk atas
    NetPortTxt = SetAmbil("porttempatan", 8180)              'port tempatan
    
    PriPmTxt(0) = SetAmbil("harga", 0.03)
    PriPmTxt(1) = SetAmbil("hargaex", 0)
    PriChk1.Value = SetAmbil("roundup", Checked)
    
    OvhTxt(0) = SetAmbil("sewa", 900)
    OvhTxt(1) = SetAmbil("bil", 230)
    OvhTxt(2) = SetAmbil("billain", 90)
    GwsOpt1(0).Value = SetAmbil("logaktiviti", Checked)
    GwsOpt1(1).Value = SetAmbil("tukarharga", Unchecked)
    
    acsTime = SetAmbil("autocloses", "12:30:00 AM")
    OverSesTxt(0) = Hour(acsTime)
    OverSesTxt(1) = Minute(acsTime)
    OverCmb1.Text = Right(acsTime, 2)
    If OverSesTxt(0) = 0 Then OverSesTxt(0) = 12
    
    If TxtNama <> "" Then TxtNama.Enabled = False Else TxtNama = "demo"         'Jika nama telah didaftar
    If TxtNombor <> "" Then TxtNombor.Enabled = False Else TxtNombor = "demo"   'disable kan textbox itu
    
    For n = 0 To uSDBe.DataCount("skema") - 1
        sk = uSDBe.DataGet("skema", "skema", n)
        hg = uSDBe.DataGet("skema", "harga", n)
        Set lItm = PriLV1.ListItems.Add(, , sk)
        lItm.SubItems(1) = hg
    Next n
    
    'loading workers
    For e = 0 To uSDBe.DataCount("pekerja-list") - 1
        strNm = uSDBe.DataGet("pekerja-list", "nama", e)
        Set lItm = Lv1.ListItems.Add(, , strNm, , "user")
        lItm.SubItems(1) = uSDBe.DataGet("pekerja-list", "gaji", e)
        lItm.SubItems(2) = uSDBe.DataGet("pekerja-list", "nick", e)
        lItm.SubItems(3) = uSDBe.DataGet("pekerja-list", "password", e)
        lItm.SubItems(4) = uSDBe.DataGet("pekerja-list", "akses", e)
        lItm.Tag = strNm
    Next e

Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, CbMsgWarn
End Sub


Private Sub Lv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strAccess As String
    strAccess = Item.SubItems(4)
    For d = 1 To 3
        Opt1(d - 1).Value = 0
        If Mid(strAccess, d, 1) = "1" Then Opt1(d - 1).Value = 1
    Next d
End Sub


Private Sub PriBtn_Click(Index As Integer)
    Dim lItm As ListItem
    Dim itmfind As ListItem
    
    Select Case ButtonIndex
    Case 0
        If PriTxt1(0) = "" Then Exit Sub
        If PriTxt1(1) = "" Then Exit Sub
        
        Set itmfind = PriLV1.FindItem(PriTxt1(0))
        If itmfind Is Nothing Then
            Set lItm = PriLV1.ListItems.Add(, , PriTxt1(0))
            lItm.SubItems(1) = PriTxt1(1)
            'tambah dalam database
            uSDBe.DataSave "skema", "skema", PriTxt1(0), True, False
            uSDBe.DataSave "Skema", "harga", PriTxt1(1), False, True
        Else
            MsgBox MB(1), vbOKOnly, CbMsgWarn: Exit Sub
        End If
    
        PriTxt1(0) = ""
        PriTxt1(1) = ""
    Case 1
        If PriLV1.ListItems.Count = 0 Then Exit Sub
        If PriLV1.SelectedItem.Text = "" Then Exit Sub
        
        nm = PriLV1.SelectedItem.Text
        lret = MsgBox(MB(2) & " " & nm & " ?", vbOKCancel, CbMsgWarn)
        If lret = vbCancel Then Exit Sub
        PriLV1.ListItems.Remove (Lv1.SelectedItem.Index)
        uSDBe.DataRemove "skema", "skema", nm
    End Select
End Sub

Private Sub RegBtn_Click()
    Dim sRet As String
    
    If CbDrvStr = "" Then CbDrvStr = "a:"
    If Trim(TxtNama) = "" And Trim(TxtNombor) = "" Then GoTo Register
    If LCase(TxtNama) = "demo" And LCase(TxtNombor) = "demo" Then GoTo Register
    
    If TxtNama <> "" And TxtNombor <> "" Then
        sRet = MsgBox(MB(5), vbOKCancel + vbInformation, CbMsgWarn)
        If sRet = vbOK Then
            If CreateDiskKey(TxtNama, TxtNombor, "a:") = True Then
                SetSimpan "demo", True
                SetSimpan "demoday", 10
                SetSimpan "namadaftar", "demo"
                SetSimpan "nombordaftar", "demo"
                CbDemoMode = True
                DemoMode = True
            End If
        End If
    End If
Exit Sub

Register:
    If ValidateDisk(CbDrvStr) = True Then
        TxtNama = GetName(CbDrvStr)
        TxtNombor = GetKey(CbDrvStr)
        If InitReg = True Then
            SetSimpan "namadaftar", TxtNama
            SetSimpan "nombordaftar", TxtNombor
            DemoMode = False
            CbDemoMode = False
            SetSimpan "demo", CStr(DemoMode)
            MsgBox MB(6), vbOKOnly, "CafeBonzer"
        End If
    End If
End Sub

Private Sub SetBtn_Click(Index As Integer)
    DemoMode = False
    
    Select Case Index
        Case 0
            If LCase(TxtNama) = "demo" And LCase(TxtNombor) = "demo" Then DemoMode = True
            '************************************
            '* simpan pada data yang program
            '* telah di buka untuk pertama kali
            '************************************
            If InitReg = False And SetAmbil("pertamakali") <> "tidak" Then
                SetSimpan "pertamakali", "ya"
                FrmSet.Hide
                Keluar False
                End
            End If
            
            'FrmSet.Hide
            Call CloseFrm(FrmSet)
            FrmMain.Show
            
        Case 1
            '************************************
            '* cek nombor pendaftaran dahulu
            '* sebelum memulakan segalanya
            '************************************
            If LCase(TxtNama) = "demo" And LCase(TxtNombor) = "demo" Then DemoMode = True
            If InitReg = False And DemoMode = False Then
                MsgBox MB(7), vbOKOnly, "Nombor ID"
                Exit Sub
            End If
            
            '************************************
            '* Periksa katalaluan utama dan client
            '************************************
            If MpassTxt(0) = "" Or MpassTxt(1) = "" Then
                MsgBox MB(8), vbInformation, CbMsgWarn
                MpassTxt(0).SetFocus
                Exit Sub
            ElseIf MpassTxt(1) <> MpassTxt(2) Then
                MpassTxt(1).SetFocus
                MsgBox MB(9), vbInformation, CbMsgWarn
                Exit Sub
            End If
            If NetPassTxt(0) <> NetPassTxt(1) Then
                NetPassTxt(0).SetFocus
                MsgBox MB(9), vbInformation, CbMsgWarn
                Exit Sub
            End If
            
            '************************************
            '* Periksa txtbox lain bagi kesalahan
            '************************************
            If NetPortTxt = "" Then NetPortTxt = 56266
            If PriPmTxt(0) = "" Or IsNumeric(Text6) = False Then Text6.Text = 0.05
            If PriPmTxt(1) = "" Or IsNumeric(Text7) = False Then Text7.Text = 0
            If OvhTxt(0) = "" Then OvhTxt(0) = 800
            If OvhTxt(1) = "" Then OvhTxt(1) = 350
            If OvhTxt(2) = "" Then OvhTxt(2) = 20
            
            '************************************
            '* Saving all settings
            '************************************
            SetSimpan "mu", MpassTxt(0)            'simpan master username..
            SetSimpan "mp", MpassTxt(1)           'simpan master password
            
            SetSimpan "namadaftar", TxtNama     'simpan no. daftar
            SetSimpan "nombordaftar", TxtNombor 'simpan nama daftar
            SetSimpan "namacc", InfoTxt(0)           'simpan nama cc
            SetSimpan "emailpengguna", InfoTxt(1)    'email pengguna
            SetSimpan "tajukatas", InfoTxt(2)        'simpan tajuk atas
            
            SetSimpan "porttempatan", NetPortTxt     'simpan no. port tempatan
            SetSimpan "netcpwd", NetPassTxt(0)
            
            SetSimpan "sewa", OvhTxt(0)
            SetSimpan "bil", OvhTxt(1)
            SetSimpan "billain", OvhTxt(2)
            SetSimpan "autocloses", (OverSesTxt(0) & ":" & OverSesTxt(1) & ":00 " & OverCmb1)
            
            SetSimpan "harga", PriPmTxt(0)            'simpan harga per/minit
            SetSimpan "hargaex", PriPmTxt(1)
            SetSimpan "roundup", PriChk1.Value
            
            SetSimpan "logaktiviti", GwsOpt1(0).Value
            SetSimpan "tukarharga", GwsOpt1(1).Value
            
            
            SetSimpan "demo", CStr(DemoMode)
            If SetAmbil("demoday") = "" Then SetSimpan "demoday", 1
            SetSimpan "demodate", Tarikh
            'tanda, bahawa pertamakali=tidak
            SetSimpan "pertamakali", "tidak"
            
            'unloadkan frmset dan menunjukkan frmmain(form utama)
            'FrmSet.Hide
            FrmMain.Caption = "CafeBonzer - " & Text5
            If SetAmbil("demo") = "True" Then FrmMain.Caption = FrmMain.Caption & " UNREGISTERED": CbDemoMode = True
            'Unload FrmSet: Set FrmSet = Nothing
            Call CloseFrm(FrmSet)
    End Select

Exit Sub
EnterNumeric:
    MsgBox MB(4), vbOKOnly, "CafeBonzer"
End Sub


