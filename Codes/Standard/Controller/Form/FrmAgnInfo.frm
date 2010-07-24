VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmAgnInfo 
   Caption         =   "Agent Information"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      ScaleHeight     =   4605
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   0
      Width           =   10170
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
         ItemData        =   "FrmAgnInfo.frx":0000
         Left            =   45
         List            =   "FrmAgnInfo.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSComctlLib.ProgressBar DynaPbar 
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   405
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
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
End
Attribute VB_Name = "FrmAgnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
