VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmSysSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CafeBonzer - Configuration"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CafeBonzer.Label3D MainLbl 
      Height          =   225
      Index           =   1
      Left            =   705
      TabIndex        =   86
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
      Caption         =   "Copyright 1996-2004 Nematix Technology"
   End
   Begin CafeBonzer.Label3D MainLbl 
      Height          =   210
      Index           =   0
      Left            =   690
      TabIndex        =   85
      Top             =   6735
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "CafeBonzer v2.0 Beta"
   End
   Begin TabDlg.SSTab MainTab 
      Height          =   6600
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   11642
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   7
      TabHeight       =   529
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmSet.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GenPassFrame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GenInfoFrame"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GenNetFrame"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GenRegFrame"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Financial"
      TabPicture(1)   =   "FrmSet.frx":6DEC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAccount"
      Tab(1).Control(1)=   "FramePricing"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Employee"
      TabPicture(2)   =   "FrmSet.frx":7386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "EmpSetChk"
      Tab(2).Control(1)=   "EmpChkSec(8)"
      Tab(2).Control(2)=   "EmpChkSec(7)"
      Tab(2).Control(3)=   "EmpChkSec(6)"
      Tab(2).Control(4)=   "EmpChkSec(5)"
      Tab(2).Control(5)=   "EmpChkSec(4)"
      Tab(2).Control(6)=   "EmpChkSec(3)"
      Tab(2).Control(7)=   "EmpTxt(0)"
      Tab(2).Control(8)=   "EmpTxt(1)"
      Tab(2).Control(9)=   "EmpTxt(2)"
      Tab(2).Control(10)=   "EmpTxt(3)"
      Tab(2).Control(11)=   "EmpChkSec(2)"
      Tab(2).Control(12)=   "EmpChkSec(1)"
      Tab(2).Control(13)=   "EmpChkSec(0)"
      Tab(2).Control(14)=   "EmpListView"
      Tab(2).Control(15)=   "EmpBtn(0)"
      Tab(2).Control(16)=   "EmpBtn(1)"
      Tab(2).Control(17)=   "EmpBtn(2)"
      Tab(2).Control(18)=   "EmpLbl(0)"
      Tab(2).Control(19)=   "EmpLbl(1)"
      Tab(2).Control(20)=   "EmpLbl(2)"
      Tab(2).Control(21)=   "EmpLbl(3)"
      Tab(2).Control(22)=   "EmpHdr(1)"
      Tab(2).Control(23)=   "EmpHdr(0)"
      Tab(2).ControlCount=   24
      Begin VB.CheckBox EmpSetChk 
         Appearance      =   0  'Flat
         Caption         =   "Log workers activities."
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -74820
         TabIndex        =   65
         Top             =   5670
         Width           =   2625
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Allow Change Price."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   -70305
         TabIndex        =   67
         Tag             =   "256"
         Top             =   6060
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Allow Cancel Client."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   -70305
         TabIndex        =   66
         Tag             =   "128"
         Top             =   5775
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Allow Unlock Client."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   -70305
         TabIndex        =   64
         Tag             =   "64"
         Top             =   5490
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Open Security Log."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   -70305
         TabIndex        =   63
         Tag             =   "32"
         Top             =   5115
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Open Agent Manager."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   -70305
         TabIndex        =   61
         Tag             =   "16"
         Top             =   4839
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Open Console."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   -70305
         TabIndex        =   58
         Tag             =   "8"
         Top             =   4563
         Width           =   2730
      End
      Begin VB.TextBox EmpTxt 
         Height          =   315
         Index           =   0
         Left            =   -73275
         TabIndex        =   43
         ToolTipText     =   "Enter worker name."
         Top             =   3735
         Width           =   2355
      End
      Begin VB.TextBox EmpTxt 
         Height          =   315
         Index           =   1
         Left            =   -73275
         TabIndex        =   47
         ToolTipText     =   "Enter the worker nick name."
         Top             =   4125
         Width           =   2355
      End
      Begin VB.TextBox EmpTxt 
         Height          =   315
         Index           =   2
         Left            =   -73275
         TabIndex        =   57
         ToolTipText     =   "Please enter the monthly salary."
         Top             =   4530
         Width           =   2355
      End
      Begin VB.TextBox EmpTxt 
         Height          =   315
         Index           =   3
         Left            =   -73275
         TabIndex        =   62
         ToolTipText     =   "Enter the password for this worker."
         Top             =   4935
         Width           =   2355
      End
      Begin VB.Frame GenRegFrame 
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
         TabIndex        =   35
         Top             =   390
         Width           =   3720
         Begin VB.TextBox GenRegName 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
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
            TabIndex        =   37
            Top             =   375
            Width           =   2220
         End
         Begin VB.TextBox GenRegNum 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
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
            TabIndex        =   39
            Top             =   870
            Width           =   2220
         End
         Begin VB.Label RegLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   36
            Top             =   420
            Width           =   630
         End
         Begin VB.Label RegLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liscence  :"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   38
            Top             =   915
            Width           =   915
         End
      End
      Begin VB.Frame GenNetFrame 
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
         Height          =   6015
         Left            =   4020
         TabIndex        =   75
         Top             =   390
         Width           =   4035
         Begin VB.TextBox GetNetPortTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2055
            MaxLength       =   5
            TabIndex        =   77
            ToolTipText     =   "Communication port between the server and the client, 56266 is the default value."
            Top             =   270
            Width           =   1830
         End
         Begin VB.TextBox GenNetPassTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   79
            ToolTipText     =   "Default password for clients, if this field left empty, it will be set same as master password."
            Top             =   780
            Width           =   1830
         End
         Begin VB.TextBox GenNetPassTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   81
            ToolTipText     =   "Communication port between the server and the client, 56266 is the default value."
            Top             =   1305
            Width           =   1830
         End
         Begin VB.Label GenNetLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Port :"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   76
            Top             =   315
            Width           =   975
         End
         Begin VB.Label GenNetLbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Client Password :"
            Height          =   420
            Index           =   1
            Left            =   210
            TabIndex        =   78
            Top             =   750
            Width           =   1380
         End
         Begin VB.Label GenNetLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retype Password :"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   80
            Top             =   1350
            Width           =   1605
         End
      End
      Begin VB.Frame GenInfoFrame 
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
         TabIndex        =   50
         Top             =   1980
         Width           =   3735
         Begin VB.TextBox GenInfoTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   255
            TabIndex        =   54
            Top             =   1260
            Width           =   3300
         End
         Begin VB.TextBox GenInfoTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   255
            TabIndex        =   52
            Top             =   555
            Width           =   3300
         End
         Begin VB.TextBox GenInfoTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   255
            TabIndex        =   56
            Top             =   1965
            Width           =   3300
         End
         Begin VB.Label GenInfoLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail :"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   53
            Top             =   990
            Width           =   675
         End
         Begin VB.Label GenInfoLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cybercafes Name :"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   51
            Top             =   285
            Width           =   1665
         End
         Begin VB.Label GenInfoLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motto :"
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   55
            Top             =   1695
            Width           =   600
         End
      End
      Begin VB.Frame FrameAccount 
         Caption         =   "Account Setting :"
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
         TabIndex        =   1
         Top             =   390
         Width           =   3660
         Begin VB.ComboBox FinSesCbDay 
            Height          =   315
            ItemData        =   "FrmSet.frx":7920
            Left            =   2010
            List            =   "FrmSet.frx":792A
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   5130
            Width           =   1515
         End
         Begin VB.ComboBox FinCmb 
            Height          =   315
            ItemData        =   "FrmSet.frx":7945
            Left            =   1440
            List            =   "FrmSet.frx":796A
            TabIndex        =   9
            Top             =   3825
            Width           =   2055
         End
         Begin VB.TextBox FinTxt 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   7
            Top             =   3435
            Width           =   2055
         End
         Begin VB.TextBox FinTxt 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   5
            Top             =   3045
            Width           =   2055
         End
         Begin VB.TextBox FinSesTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   1590
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "12"
            Top             =   5580
            Width           =   450
         End
         Begin VB.TextBox FinSesTxt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   2250
            MaxLength       =   2
            TabIndex        =   19
            Text            =   "30"
            Top             =   5580
            Width           =   450
         End
         Begin VB.ComboBox FinSesCmb 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "FrmSet.frx":79CB
            Left            =   2820
            List            =   "FrmSet.frx":79D5
            TabIndex        =   20
            Text            =   "AM"
            Top             =   5580
            Width           =   690
         End
         Begin MSComctlLib.ListView FinListView 
            Height          =   2250
            Left            =   150
            TabIndex        =   3
            Top             =   645
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   3969
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Overhead"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Filter"
               Object.Width           =   1411
            EndProperty
         End
         Begin CafeBonzer.XpButton FinBtn 
            Height          =   360
            Index           =   1
            Left            =   2640
            TabIndex        =   11
            Top             =   4260
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
            MICON           =   "FrmSet.frx":79E1
            PICN            =   "FrmSet.frx":79FD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton FinBtn 
            Height          =   360
            Index           =   0
            Left            =   2205
            TabIndex        =   10
            Top             =   4260
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
            MICON           =   "FrmSet.frx":7F97
            PICN            =   "FrmSet.frx":7FB3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton FinBtn 
            Height          =   360
            Index           =   2
            Left            =   3075
            TabIndex        =   12
            Top             =   4260
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
            MICON           =   "FrmSet.frx":834D
            PICN            =   "FrmSet.frx":8369
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label FinLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Session Close :"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   18
            Top             =   5595
            Width           =   1335
         End
         Begin VB.Label FinLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Month Filter :"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   8
            Top             =   3855
            Width           =   1125
         End
         Begin VB.Label FinLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value :"
            Height          =   195
            Index           =   1
            Left            =   735
            TabIndex        =   6
            Top             =   3450
            Width           =   570
         End
         Begin VB.Label FinLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overhead :"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   4
            Top             =   3075
            Width           =   975
         End
         Begin VB.Label FinHdr 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Monthly Fixed Overhead"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   255
            Width           =   3645
         End
         Begin VB.Label FinHdr 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Account Settings"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   4785
            Width           =   3645
         End
         Begin VB.Label FinLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Session Close Day :"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   14
            Top             =   5190
            Width           =   1740
         End
         Begin VB.Label FinLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   195
            Index           =   4
            Left            =   2100
            TabIndex        =   17
            Top             =   5595
            Width           =   75
         End
      End
      Begin VB.Frame GenPassFrame 
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
         TabIndex        =   68
         Top             =   4605
         Width           =   3720
         Begin VB.TextBox GenPassTxt 
            Alignment       =   2  'Center
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
            TabIndex        =   70
            Top             =   285
            Width           =   1965
         End
         Begin VB.TextBox GenPassTxt 
            Alignment       =   2  'Center
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
            TabIndex        =   72
            Top             =   765
            Width           =   1965
         End
         Begin VB.TextBox GenPassTxt 
            Alignment       =   2  'Center
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
            TabIndex        =   74
            Top             =   1245
            Width           =   1965
         End
         Begin VB.Label GenPassLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username :"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   69
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label GenPassLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   71
            Top             =   795
            Width           =   945
         End
         Begin VB.Label GenPassLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retype :"
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   73
            Top             =   1275
            Width           =   735
         End
      End
      Begin VB.Frame FramePricing 
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
         TabIndex        =   21
         Top             =   390
         Width           =   4125
         Begin VB.TextBox PriTxt 
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   29
            Top             =   3840
            Width           =   1815
         End
         Begin VB.CheckBox PriChk 
            Appearance      =   0  'Flat
            Caption         =   "Round Up Price"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   135
            TabIndex        =   34
            Top             =   4890
            Width           =   2220
         End
         Begin VB.TextBox PriTxt 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   25
            Top             =   3045
            Width           =   1815
         End
         Begin VB.TextBox PriTxt 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   27
            Top             =   3435
            Width           =   1815
         End
         Begin CafeBonzer.XpButton PriBtn 
            Height          =   360
            Index           =   0
            Left            =   3120
            TabIndex        =   30
            Top             =   4275
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
            MICON           =   "FrmSet.frx":8903
            PICN            =   "FrmSet.frx":891F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.Line3D PriLine 
            Height          =   45
            Left            =   60
            TabIndex        =   33
            Top             =   4755
            Width           =   3990
            _ExtentX        =   7038
            _ExtentY        =   79
            horizon         =   -1  'True
         End
         Begin MSComctlLib.ListView PriListView 
            Height          =   2280
            Left            =   150
            TabIndex        =   23
            Top             =   630
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   4022
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Scheme"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Price\Hour"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Extra"
               Object.Width           =   1411
            EndProperty
         End
         Begin CafeBonzer.XpButton PriBtn 
            Height          =   360
            Index           =   1
            Left            =   2685
            TabIndex        =   31
            Top             =   4275
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
            MICON           =   "FrmSet.frx":8EB9
            PICN            =   "FrmSet.frx":8ED5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton PriBtn 
            Height          =   360
            Index           =   2
            Left            =   3555
            TabIndex        =   32
            Top             =   4275
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
            MICON           =   "FrmSet.frx":926F
            PICN            =   "FrmSet.frx":928B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label PriLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Initial Charges :"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   28
            Top             =   3855
            Width           =   1395
         End
         Begin VB.Label PriHdr1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Normal Pricing"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   255
            Width           =   4110
         End
         Begin VB.Label PriLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scheme Name :"
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   24
            Top             =   3075
            Width           =   1380
         End
         Begin VB.Label PriLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price per Minute :"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   26
            Top             =   3450
            Width           =   1515
         End
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Open Statistic."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   -70305
         TabIndex        =   49
         Tag             =   "4"
         Top             =   4287
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Open Configuration."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   -70305
         TabIndex        =   46
         Tag             =   "2"
         Top             =   4011
         Width           =   2730
      End
      Begin VB.CheckBox EmpChkSec 
         Appearance      =   0  'Flat
         Caption         =   "Open External Modules."
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   -70305
         TabIndex        =   44
         Tag             =   "1"
         Top             =   3735
         Width           =   2625
      End
      Begin MSComctlLib.ListView EmpListView 
         Height          =   2715
         Left            =   -74865
         TabIndex        =   40
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
         Left            =   -73965
         TabIndex        =   82
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
         MICON           =   "FrmSet.frx":9825
         PICN            =   "FrmSet.frx":9841
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
         Left            =   -74400
         TabIndex        =   83
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
         MICON           =   "FrmSet.frx":9DDB
         PICN            =   "FrmSet.frx":9DF7
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
         Index           =   2
         Left            =   -74835
         TabIndex        =   84
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
         MICON           =   "FrmSet.frx":A191
         PICN            =   "FrmSet.frx":A1AD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label EmpLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Worker Name :"
         Height          =   300
         Index           =   0
         Left            =   -74745
         TabIndex        =   45
         Top             =   3780
         Width           =   1440
      End
      Begin VB.Label EmpLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName :"
         Height          =   300
         Index           =   1
         Left            =   -74745
         TabIndex        =   48
         Top             =   4170
         Width           =   1440
      End
      Begin VB.Label EmpLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary :"
         Height          =   300
         Index           =   2
         Left            =   -74745
         TabIndex        =   59
         Top             =   4965
         Width           =   1440
      End
      Begin VB.Label EmpLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         Height          =   300
         Index           =   3
         Left            =   -74745
         TabIndex        =   60
         Top             =   4575
         Width           =   1440
      End
      Begin VB.Label EmpHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Employee Information"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   -74865
         TabIndex        =   42
         Top             =   3345
         Width           =   4155
      End
      Begin VB.Label EmpHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Employee Security Settings"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   -70590
         TabIndex        =   41
         Top             =   3345
         Width           =   3660
      End
   End
   Begin CafeBonzer.XpButton MainBtn 
      Height          =   435
      Index           =   1
      Left            =   7725
      TabIndex        =   88
      ToolTipText     =   "Save settings and exit."
      Top             =   6855
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
      MICON           =   "FrmSet.frx":A747
      PICN            =   "FrmSet.frx":A763
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton MainBtn 
      Height          =   435
      Index           =   0
      Left            =   7170
      TabIndex        =   87
      ToolTipText     =   "Cancel all settings and exit."
      Top             =   6855
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
      MICON           =   "FrmSet.frx":ACFD
      PICN            =   "FrmSet.frx":AD19
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image MainImg 
      Height          =   480
      Left            =   90
      Picture         =   "FrmSet.frx":B2B3
      Top             =   6750
      Width           =   480
   End
End
Attribute VB_Name = "FrmSysSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmSysSet
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'
' New Setting Redefined
'
'   RegName         = Register Name\Owner
'   RegNumber       = Register Number\Key
'   GenAdminName    = Admin Name
'   GenAdminPass    = Admin Password
'   GenOrgName      = Organization Name
'   GenOrgEmail     = Organization Email
'   GenOrgMoto      = Organization Motto
'   NetPortLocal    = Main Local Port
'   FinPriceRound   = Round Up a Price
'   FinSessionClose = Close Session Time
'   SecLogUser      = Log User Activity
'
'   AppFirstTime    = First Time Flag
'   AppDemoDay      = Total Use Demo Day
'   AppDemoDate     =
'
'   AppMainNote     = Main Note
'   UiToolBar       = Main Toolbar Setting
'   UiDockBar       = Main DockBar Setting
'
'

Private Sub Form_Initialize()
    NumOnly GetNetPortTxt
    NumOnly FinSesTxt(0)
    NumOnly FinSesTxt(1)
End Sub

Private Sub Form_Load()
On Error GoTo ErrInt
    Dim StrTmpSessClose As String
    
    GenRegName = SettingGet("RegName", "Demo")
    GenRegNum = SettingGet("RegNumber", "Demo")
    
    GenPassTxt(0) = SetGetDb("GenAdminName", "admin")
    GenPassTxt(1) = SetGetDb("GenAdminPass")
    GenPassTxt(2) = GenPassTxt(1)
    
    GenInfoTxt(0) = SetGetDb("GenOrgName")
    GenInfoTxt(1) = SetGetDb("GenOrgEmail")
    GenInfoTxt(2) = SetGetDb("GenOrgMoto")
    GetNetPortTxt = SetGetDb("NetPortLocal", 8180)
    GenNetPassTxt(0) = SetGetDb("NetDefaultPass")
    GenNetPassTxt(1) = GenNetPassTxt(0)
    
    PriChk.Value = SetGetDb("FinPriceRound", Checked)
    EmpSetChk.Value = SetGetDb("SecLogUser", Checked)

    FinSesCbDay.ListIndex = SetGetDb("FinSessionDay", 0)
    StrTmpSessClose = SetGetDb("FinSessionClose", "12:30:00 AM")
    FinSesTxt(0) = Hour(StrTmpSessClose)
    FinSesTxt(1) = Minute(StrTmpSessClose)
    FinSesCmb.Text = Right(StrTmpSessClose, 2)
    If FinSesTxt(0) = 0 Then FinSesTxt(0) = 12
    
    Call FinOverheadAction(0)
    Call FinSchemeAction(0)
    Call EmployeeAction(0)
Exit Sub

ErrInt:
    MsgBox Err.Description, vbExclamation, CbMsgWarn
End Sub

Private Sub MainBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            '[ First Time Flag ]'
            If SettingGet("AppFirstTime") <> "NoFirst" Then
                SettingSave "AppFirstTime", "YesFirst"
                FrmSysSet.Hide
                End
            End If
            Call FormClose(FrmSysSet)
            FrmMain.Show
            
        Case 1
            '[ Check For Data Validity ]'
            If CtlCheckNull(GenPassTxt(0), ST(0, 1)) = False Then Exit Sub
            If CtlCheckNull(GenPassTxt(1), ST(0, 1)) = False Then Exit Sub
            If CtlCheckMatch(GenPassTxt(1), GenPassTxt(2), ST(0, 2)) = False Then Exit Sub
            If CtlCheckMatch(GenNetPassTxt(0), GenNetPassTxt(1), ST(0, 2)) = False Then Exit Sub
            If GetNetPortTxt = "" Then GetNetPortTxt = 8180

            '[ Saving All Settings ]'
            SetSaveDb "GenAdminName", GenPassTxt(0)
            SetSaveDb "GenAdminPass", GenPassTxt(1)
            SetSaveDb "GenOrgName", GenInfoTxt(0)
            SetSaveDb "GenOrgEmail", GenInfoTxt(1)
            SetSaveDb "GenOrgMoto", GenInfoTxt(2)
            SetSaveDb "NetPortLocal", GetNetPortTxt
            SetSaveDb "NetDefaultPass", GenNetPassTxt(0)
            
            SetSaveDb "FinSessionDay", FinSesCbDay.ListIndex
            SetSaveDb "FinSessionClose", (FinSesTxt(0) & ":" & FinSesTxt(1) & ":00 " & FinSesCmb)
            SetSaveDb "FinPriceRound", PriChk.Value
            SetSaveDb "SecLogUser", EmpSetChk.Value
            
            SettingSave "RegName", GenRegName
            SettingSave "RegNumber", GenRegNum
            SettingSave "AppFirstTime", "NoFirst"

            FrmMain.Caption = "CafeBonzer - " & GenInfoTxt(2)
            Call FormClose(FrmSysSet)
    End Select
End Sub

Private Sub FinBtn_Click(Index As Integer)
    Call FinOverheadAction(Index + 1)
End Sub

Private Sub EmpBtn_Click(Index As Integer)
    Call EmployeeAction(Index + 1)
End Sub

Private Sub PriBtn_Click(Index As Integer)
    Call FinSchemeAction(Index + 1)
End Sub

Private Sub FinListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    FinTxt(0) = Item.Text
    FinTxt(1) = Item.SubItems(1)
    FinCmb = Item.SubItems(2)
End Sub

Private Sub FinSesCbDay_Click()
    Dim LngSesDay As Long
    LngSesDay = FinSesCbDay.ListIndex
    FinSesTxt(0) = Switch(LngSesDay = 0, "12", LngSesDay = 1, "11")
    FinSesCmb = Switch(LngSesDay = 0, "AM", LngSesDay = 1, "PM")
End Sub

Private Sub PriListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PriTxt(0) = Item.Text
    PriTxt(1) = Item.SubItems(1)
    PriTxt(2) = Item.SubItems(2)
End Sub

Private Sub EmpListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim LngTmpUserAccess As Long, LngIdxA As Long

    LngTmpUserAccess = Item.SubItems(4)
    EmpTxt(0) = Item.Text
    EmpTxt(1) = Item.SubItems(2)
    EmpTxt(2) = Item.SubItems(3)
    EmpTxt(3) = Item.SubItems(1)
    
    For LngIdxA = 0 To LngTotalAccessCode - 1
        EmpChkSec(LngIdxA).Value = 0
        If SecAccessRequest(2 ^ LngIdxA, LngTmpUserAccess) = True Then EmpChkSec(LngIdxA).Value = 1
    Next LngIdxA
End Sub

Private Sub FinOverheadAction(Action As Long)
    Dim CListItem As ListItem, LngIdxA As Long, StrOverhead As String, strValue As String
    Dim StrFilter As String, LngRet As Long, CRset As Recordset
    ' Load Overhead = 0
    ' Del Overhead  = 1
    ' Add Overhead  = 2
    ' Save Overhead = 3
    
    StrOverhead = FinTxt(0)
    strValue = FinTxt(1)
    StrFilter = FinCmb.Text
        
    Select Case Action
    Case 0
        For LngIdxA = 0 To CDataSe.DataCount("FinanceOverhead") - 1
            StrOverhead = CDataSe.DataGet("FinanceOverhead", "Overhead", LngIdxA)
            Set CListItem = FinListView.ListItems.Add(, , StrOverhead)
            CListItem.SubItems(1) = CDataSe.DataGet("FinanceOverhead", "Value", LngIdxA)
            CListItem.SubItems(2) = CDataSe.DataGet("FinanceOverhead", "DisOccurMonth", LngIdxA)
        Next LngIdxA
    Case 1
        If FinListView.ListItems.Count = 0 Then Exit Sub
        If FinListView.SelectedItem.Text = "" Then Exit Sub
        StrOverhead = FinListView.SelectedItem.Text
        LngRet = MsgBox(VS(0, 0) & " " & StrOverhead & " ?", vbOKCancel, CbMsgWarn)
        If LngRet = vbOK Then
            Set CRset = CDataS.OpenRecordset("FinanceOverhead", dbOpenDynaset)
            CRset.FindFirst "Overhead = '" + StrOverhead + "'"
            If CRset.NoMatch = False Then
                CRset.Delete
                FinListView.ListItems.Remove (FinListView.SelectedItem.Index)
            End If
            CRset.Close
        End If
    Case 2
        If StrOverhead = "" Or strValue = "" Then Exit Sub
        Set ClistItemFind = FinListView.FindItem(StrOverhead)
        If ClistItemFind Is Nothing Then
            CDataSe.DataSave "FinanceOverhead", "Overhead", StrOverhead, True, False
            CDataSe.DataSave "FinanceOverhead", "Value", strValue, False, False
            CDataSe.DataSave "FinanceOverhead", "DisOccurMonth", StrFilter, False, True
            Set CListItem = FinListView.ListItems.Add(, , StrOverhead)
            CListItem.SubItems(1) = strValue
            CListItem.SubItems(2) = StrFilter
        Else
            MsgBox ST(0, 4), vbOKOnly, CbMsgWarn
            Exit Sub
        End If
    Case 3
        Set CRset = CDataS.OpenRecordset("FinanceOverhead", dbOpenDynaset)
        CRset.FindFirst "Overhead = '" + StrOverhead + "'"
        If CRset.NoMatch = False Then
            CRset.Edit
            CRset.Fields("Overhead").Value = StrOverhead
            CRset.Fields("Value").Value = strValue
            CRset.Fields("DisOccurMonth").Value = StrFilter
            CRset.Update
            FinListView.SelectedItem.Text = StrOverhead
            FinListView.SelectedItem.SubItems(1) = strValue
            FinListView.SelectedItem.SubItems(2) = StrFilter
        End If
        CRset.Close
    End Select
End Sub

Private Sub FinSchemeAction(Action As Long)
    Dim CListItem As ListItem, ClistItemFind As ListItem, StrSchemeIndex As String
    Dim StrName As String, StrPrice As String, StrExtra As String
    ' Load Scheme   = 0
    ' Add Scheme    = 1
    ' Del Scheme    = 2
    ' Save Scheme   = 3
    
    StrName = PriTxt(0)
    StrPrice = PriTxt(1)
    StrExtra = PriTxt(2)
    
    Select Case Action
    Case 0
        For LngIdxA = 0 To CDataSe.DataCount("PriceScheme") - 1
            StrName = CDataSe.DataGet("PriceScheme", "Scheme", LngIdxA)
            Set CListItem = PriListView.ListItems.Add(, , StrName)
            CListItem.SubItems(1) = CDataSe.DataGet("PriceScheme", "Price", LngIdxA)
            CListItem.SubItems(2) = CDataSe.DataGet("PriceScheme", "Extra", LngIdxA)
        Next LngIdxA
    Case 1
        If StrName = "" Or StrPrice = "" Then Exit Sub
        If StrExtra = "" Then StrExtra = 0
        Set ClistItemFind = PriListView.FindItem(StrName)
        If ClistItemFind Is Nothing Then
            CDataSe.DataSave "PriceScheme", "Scheme", StrName, True, False
            CDataSe.DataSave "PriceScheme", "Price", StrPrice, False, False
            CDataSe.DataSave "PriceScheme", "Extra", StrExtra, False, True
            Set CListItem = PriListView.ListItems.Add(, , StrName)
            CListItem.SubItems(1) = StrPrice
            CListItem.SubItems(2) = StrExtra
        Else
            MsgBox ST(0, 3), vbOKOnly, CbMsgWarn
            Exit Sub
        End If
    Case 2
        If PriListView.ListItems.Count = 0 Then Exit Sub
        If PriListView.SelectedItem.Text = "" Then Exit Sub
        StrName = PriListView.SelectedItem.Text
        If StrName = "Default" Then Exit Sub
        lret = MsgBox(VS(0, 0) & " " & StrName & " ?", vbOKCancel, CbMsgWarn)
        If lret = vbOK Then
            CDataSe.DataRemove "PriceScheme", "Scheme", StrName
            PriListView.ListItems.Remove (PriListView.SelectedItem.Index)
        End If
    Case 3
        If StrName = "" Or StrPrice = "" Then Exit Sub
        If StrExtra = "" Then StrExtra = 0
        StrSchemeIndex = PriListView.SelectedItem.Text
        CDataSe.DataEdit "PriceScheme", "Scheme", "Scheme", StrSchemeIndex, StrName, True, False
        CDataSe.DataEdit "PriceScheme", "Price", "Scheme", StrSchemeIndex, StrPrice, False, False
        CDataSe.DataEdit "PriceScheme", "Extra", "Scheme", StrSchemeIndex, StrExtra, False, True
        PriListView.SelectedItem.Text = StrName
        PriListView.SelectedItem.SubItems(1) = StrPrice
        PriListView.SelectedItem.SubItems(2) = StrExtra
    End Select
End Sub

Private Sub EmployeeAction(Action As Long)
    Dim CListItem As ListItem, LngIdxA As Long, StrName As String, LngTmpUserAccess As Long
    Dim CRset As Recordset
    ' Load Employee = 0
    ' Add Employee = 1
    ' Del Employee = 2
    ' Save Access = 3
    
    Select Case Action
    Case 0
        For LngIdxA = 0 To CDataSe.DataCount("ListEmployee") - 1
            StrName = CDataSe.DataGet("ListEmployee", "Name", LngIdxA)
            Set CListItem = EmpListView.ListItems.Add(, , StrName)
            CListItem.SubItems(1) = CDataSe.DataGet("ListEmployee", "Salary", LngIdxA)
            CListItem.SubItems(2) = CDataSe.DataGet("ListEmployee", "UserName", LngIdxA)
            CListItem.SubItems(3) = CDataSe.DataGet("ListEmployee", "Password", LngIdxA)
            CListItem.SubItems(4) = CDataSe.DataGet("ListEmployee", "Access", LngIdxA)
            CListItem.Tag = StrName
        Next LngIdxA
    Case 1
        If EmpTxt(0) = "" Then Exit Sub
        If EmpTxt(1) = "" Then Exit Sub
        If EmpTxt(2) = "" Then Exit Sub
        If EmpTxt(3) = "" Then Exit Sub
    
        Set CListItem = EmpListView.ListItems.Add(, , EmpTxt(1))
        CListItem.SubItems(1) = EmpTxt(1)
        CListItem.SubItems(2) = EmpTxt(2)
        CListItem.SubItems(3) = EmpTxt(3)
        CListItem.SubItems(4) = 1
        
        CDataSe.DataSave "ListEmployee", "Name", EmpTxt(0), True, False
        CDataSe.DataSave "ListEmployee", "UserName", EmpTxt(1), False, False
        CDataSe.DataSave "ListEmployee", "Password", EmpTxt(2), False, False
        CDataSe.DataSave "ListEmployee", "Salary", EmpTxt(3), False, False
        CDataSe.DataSave "ListEmployee", "Access", 1, False, True
    Case 2
        If EmpListView.ListItems.Count = 0 Then Exit Sub
        StrName = EmpListView.SelectedItem.Tag
        lret = MsgBox("Delete " & EmpListView.SelectedItem.Text & " ?", vbOKCancel, CbMsgWarn)
        If lret = vbCancel Then Exit Sub
        CDataSe.DataRemove "ListEmployee", "Name", StrName
        EmpListView.ListItems.Remove (EmpListView.SelectedItem.Index)
    Case 3
        If EmpListView.ListItems.Count = 0 Then Exit Sub
        For LngIdxA = 0 To LngTotalAccessCode - 1
            If EmpChkSec(LngIdxA).Value = 1 Then
                LngTmpUserAccess = LngTmpUserAccess + EmpChkSec(LngIdxA).Tag
            End If
        Next LngIdxA
        Set CRset = CDataS.OpenRecordset("ListEmployee", dbOpenDynaset)
        CRset.FindFirst "UserName = '" + EmpListView.SelectedItem.SubItems(2) + "'"
        If CRset.NoMatch = False Then
            CRset.Edit
            CRset.Fields("Name").Value = EmpTxt(0)
            CRset.Fields("UserName").Value = EmpTxt(1)
            CRset.Fields("Password").Value = EmpTxt(2)
            CRset.Fields("Salary").Value = EmpTxt(3)
            CRset.Fields("Access").Value = LngTmpUserAccess
            CRset.Update
            EmpListView.SelectedItem.Text = EmpTxt(0)
            EmpListView.SelectedItem.SubItems(1) = EmpTxt(3)
            EmpListView.SelectedItem.SubItems(2) = EmpTxt(1)
            EmpListView.SelectedItem.SubItems(3) = EmpTxt(2)
            EmpListView.SelectedItem.SubItems(4) = LngTmpUserAccess
        End If
        CRset.Close
    End Select
End Sub

' 18/08/2003 - Third annivesarry with my love, still counting
