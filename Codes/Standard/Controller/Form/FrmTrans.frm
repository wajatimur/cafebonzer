VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAgnTrans 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4065
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3885
   ClipControls    =   0   'False
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
   Icon            =   "FrmTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin CafeBonzer.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.PictureBox LgBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   3870
      TabIndex        =   2
      Top             =   0
      Width           =   3870
      Begin CafeBonzer.Label3D LgBannerTerminal 
         Height          =   300
         Left            =   1875
         TabIndex        =   3
         ToolTipText     =   "Nama Pc"
         Top             =   60
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
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
         Caption         =   "Transfer PC"
         BackColor       =   16777215
      End
      Begin VB.Image LgBannerImg 
         Height          =   960
         Left            =   75
         Picture         =   "FrmTrans.frx":000C
         Top             =   30
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView List1 
      Height          =   2160
      Left            =   60
      TabIndex        =   1
      Top             =   1305
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station"
         Object.Width           =   4233
      EndProperty
   End
   Begin CafeBonzer.XpButton BtnOk 
      Height          =   450
      Left            =   2820
      TabIndex        =   5
      ToolTipText     =   "Add new employee."
      Top             =   3540
      Width           =   990
      _ExtentX        =   1746
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmTrans.frx":0A88
      PICN            =   "FrmTrans.frx":0AA4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton BtnKo 
      Height          =   450
      Left            =   1785
      TabIndex        =   6
      ToolTipText     =   "Add new employee."
      Top             =   3540
      Width           =   990
      _ExtentX        =   1746
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmTrans.frx":103E
      PICN            =   "FrmTrans.frx":105A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label TermName 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Terminal01"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   900
      Width           =   3750
   End
End
Attribute VB_Name = "FrmAgnTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmAgnTrans
'    Project    : CafeBonzer
'
'    Description: Transfer Agent To Another Terminals
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub BtnKo_Click()
    Unload Me
End Sub

Private Sub BtnOk_Click()
    If List1.ListItems.Count = 0 Then Exit Sub
    
    If List1.SelectedItem.Text = "" Then
        FrmAgnTrans.Hide
        Unload FrmAgnTrans
    Else
        FrmSysHost.TmrAgent.Enabled = False
        FrmSysHost.TmrPing.Enabled = False
        
        Call AgentSel.AgnTransfer(List1.SelectedItem.Tag)
        
        FrmSysHost.TmrAgent.Enabled = True
        FrmSysHost.TmrPing.Enabled = True
        Unload FrmAgnTrans
    End If
End Sub

Private Sub Form_Load()
    Dim ClvMain As ListView, ClvItem As ListItem
    Dim LngIdx As Long, LngItemCnt As Long
    
    Set ClvMain = FrmMain.ListView
    List1.SmallIcons = FrmMain.ImgList16
    TermName = ClvMain.SelectedItem.Text
    
    LngItemCnt = ClvMain.ListItems.Count
    For LngIdx = 1 To LngItemCnt
        If ClvMain.ListItems(LngIdx).Text <> TermName Then
            If ClvMain.ListItems(LngIdx).SubItems(1) = VS(1, 2) Then
                Set ClvItem = List1.ListItems.Add(, , ClvMain.ListItems(LngIdx).Text, , "TerminalOnline")
                ClvItem.Tag = ClvMain.ListItems(LngIdx).Index
            End If
        End If
    Next
End Sub

