VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGrpAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Group"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGrpAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CbSnmMgr.Line3D Line3D1 
      Height          =   45
      Left            =   15
      TabIndex        =   9
      Top             =   2445
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin MSComctlLib.ListView GrpSymLv 
      Height          =   945
      Left            =   135
      TabIndex        =   2
      Top             =   1410
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   1667
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
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
   Begin VB.TextBox GrpInfoTxt 
      Height          =   330
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   150
      Width           =   2145
   End
   Begin VB.TextBox GrpInfoId 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   150
      Width           =   1005
   End
   Begin VB.TextBox GrpInfoTxt 
      Height          =   330
      Index           =   1
      Left            =   1305
      TabIndex        =   1
      Top             =   630
      Width           =   3270
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   450
      Index           =   0
      Left            =   3675
      TabIndex        =   3
      ToolTipText     =   "Accept"
      Top             =   2565
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
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
      MICON           =   "FrmGrpAdd.frx":23D2
      PICN            =   "FrmGrpAdd.frx":23EE
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
      Index           =   1
      Left            =   4155
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   2565
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
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
      MICON           =   "FrmGrpAdd.frx":2988
      PICN            =   "FrmGrpAdd.frx":29A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group :"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   195
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Symbol :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   675
      Width           =   1095
   End
End
Attribute VB_Name = "FrmGrpAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnMenu_Click(Index As Integer)
    Select Case Index
    Case 0
        Call GroupAdd(GrpInfoId, GrpInfoTxt(0), GrpInfoTxt(1), GrpSymLv.SelectedItem.Key)
    End Select
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim IntIdxA As Integer, IntIdxB As Integer
    
    Set GrpSymLv.Icons = FrmSnmMg.Iml
    
    Call LoadSymbol(GrpSymLv, True)
    GrpInfoId = GroupIdCheck
End Sub

