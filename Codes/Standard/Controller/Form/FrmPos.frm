VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPos 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4905
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   9180
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
   Icon            =   "FrmPos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4905
   ScaleWidth      =   9180
   StartUpPosition =   1  'CenterOwner
   Begin CafeBonzer.VsGuiSpinEdit SerQty 
      Height          =   435
      Left            =   1755
      TabIndex        =   23
      Top             =   3720
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   767
      ENABLED         =   0   'False
      QUANTITY        =   1
   End
   Begin CafeBonzer.Line3D Line3D3 
      Height          =   45
      Left            =   15
      TabIndex        =   21
      Top             =   4230
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   780
      TabIndex        =   19
      Top             =   3780
      Visible         =   0   'False
      Width           =   900
   End
   Begin CafeBonzer.Line3D Line3D2 
      Height          =   4140
      Left            =   4380
      TabIndex        =   18
      Top             =   750
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   7303
      horizon         =   0   'False
   End
   Begin VB.TextBox PayRcv 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1275
      TabIndex        =   0
      ToolTipText     =   "Enter Received Value."
      Top             =   4395
      Width           =   1515
   End
   Begin VB.PictureBox PayBox 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   7050
      ScaleHeight     =   735
      ScaleWidth      =   2145
      TabIndex        =   6
      Top             =   0
      Width           =   2145
      Begin VB.Label PayBal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   1125
         TabIndex        =   10
         Top             =   420
         Width           =   585
      End
      Begin VB.Label PayTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   1125
         TabIndex        =   9
         Top             =   60
         Width           =   585
      End
      Begin VB.Label PayLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   8
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label PayLbl 
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
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   2
         Left            =   390
         TabIndex        =   7
         Top             =   45
         Width           =   675
      End
   End
   Begin VB.PictureBox LgBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7050
      TabIndex        =   4
      Top             =   0
      Width           =   7050
      Begin CafeBonzer.Label3D LgBannerTerminal 
         Height          =   300
         Left            =   990
         TabIndex        =   5
         ToolTipText     =   "Nama Pc"
         Top             =   45
         Width           =   5985
         _ExtentX        =   10557
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
         Caption         =   "Point of Sales"
         BackColor       =   14737632
      End
      Begin VB.Image LgBannerImg 
         Height          =   960
         Left            =   60
         Picture         =   "FrmPos.frx":000C
         Top             =   15
         Width           =   960
      End
   End
   Begin CafeBonzer.Line3D Line3D1 
      Height          =   45
      Index           =   0
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CafeBonzer.XpButton BtnMnu 
      Height          =   585
      Index           =   0
      Left            =   7800
      TabIndex        =   1
      ToolTipText     =   "Confirm Transaction"
      Top             =   4260
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1032
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "FrmPos.frx":0A71
      PICN            =   "FrmPos.frx":0A8D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton BtnMnu 
      Height          =   585
      Index           =   1
      Left            =   8475
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   4260
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1032
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
      MICON           =   "FrmPos.frx":1027
      PICN            =   "FrmPos.frx":1043
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageCombo SerICmb 
      Height          =   330
      Left            =   1185
      TabIndex        =   12
      Top             =   855
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Locked          =   -1  'True
      Text            =   "None"
   End
   Begin MSComctlLib.ListView SerLv 
      Height          =   2370
      Left            =   75
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1305
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4180
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Stock"
         Object.Width           =   1411
      EndProperty
   End
   Begin CafeBonzer.XpButton SerBtn 
      Height          =   345
      Index           =   1
      Left            =   3795
      TabIndex        =   14
      Top             =   3765
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "FrmPos.frx":15DD
      PICN            =   "FrmPos.frx":15F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LvItems 
      Height          =   3285
      Left            =   4545
      TabIndex        =   16
      Top             =   900
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qty"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Total"
         Object.Width           =   1411
      EndProperty
   End
   Begin CafeBonzer.XpButton BtnItems 
      Height          =   345
      Index           =   0
      Left            =   4545
      TabIndex        =   17
      Top             =   4260
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
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
      MICON           =   "FrmPos.frx":1B93
      PICN            =   "FrmPos.frx":1BAF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton BtnMnu 
      Height          =   585
      Index           =   2
      Left            =   7125
      TabIndex        =   22
      ToolTipText     =   "Add To Selected User"
      Top             =   4260
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1032
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "FrmPos.frx":2149
      PICN            =   "FrmPos.frx":2165
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
      Caption         =   "Code :"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   3810
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label SerLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Category :"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   105
      TabIndex        =   15
      Top             =   885
      Width           =   1005
   End
   Begin VB.Label PayLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receive :"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4365
      Width           =   1005
   End
End
Attribute VB_Name = "FrmPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmPos
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private DblServicePaid As Double    'CalculateTotal


Private Sub Form_Load()
    LvItems.SmallIcons = FrmMain.ImgListSnm
    
    'If SecAccessRequest(TaskChangePrice) = False Then PayPrice.Locked = True
    Call LoadPosCatCB(SerICmb, FrmMain.ImgListSnm, True)
End Sub

Private Sub PayRcv_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CalculateTotal
    If KeyCode = vbKeyReturn Then BtnMnu_Click (0)
End Sub

Private Sub BtnItems_Click(Index As Integer)
    Select Case Index
    Case 0
        If LvItems.SelectedItem.Text <> VS(2, 0) Then
            LvItems.ListItems.Remove (LvItems.SelectedItem.Index)
            Call CalculateTotal
        End If
    End Select
End Sub

Private Sub BtnMnu_Click(Index As Integer)
    Select Case Index
    Case 0
        Call LogUsage
    Case 2
        '( Add to user )'
    End Select
    Unload Me
End Sub


Private Sub SerBtn_Click(Index As Integer)
    Dim TmpItem As ListItem, SelItem As ListItem, FindItm As ListItem
    Set TmpItem = SerLv.SelectedItem
    
    If SerQty > 0 Then
        Set FindItm = LvItems.FindItem(TmpItem.Text)
        If FindItm Is Nothing Then
            Set SelItem = LvItems.ListItems.Add(, TmpItem.Key, TmpItem.Text, , "ITEM")
            SelItem.SubItems(1) = SerQty
            SelItem.SubItems(2) = Format(SerQty * TmpItem.SubItems(1), "0.00")
        Else
            FindItm.SubItems(1) = SerQty
            FindItm.SubItems(2) = Format(SerQty * TmpItem.SubItems(1), "0.00")
        End If
        Call CalculateTotal
    Else
        MsgBox ST(0, 0), vbOKOnly, CbMsgWarn
        SerQty.EditValue = 1
    End If
End Sub

Private Sub SerICmb_Click()
    Dim BlnLoaded As Boolean
    
    Call SerControl(True)
    BlnLoaded = LoadPosItmLV(SerLv, SerICmb.SelectedItem.Key, FrmMain.ImgListSnm)
    If BlnLoaded = False Then SerControl (False)
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' FUNCTION
'
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Total Calculation
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub CalculateTotal()
    Dim IntIdxA As Integer, LngItemCnt As Long
    
    DblServicePaid = 0
    If LvItems.ListItems.Count > 0 Then
        LngItemCnt = LvItems.ListItems.Count
        For IntIdxA = 1 To LngItemCnt
            If LvItems.ListItems(IntIdxA).Text <> VS(2, 0) Then
                DblServicePaid = DblServicePaid + CDbl(LvItems.ListItems(IntIdxA).SubItems(2))
            End If
        Next
        PayTotal = Format(DblServicePaid, "0.00")
        BtnMnu(0).Enabled = True
        BtnMnu(2).Enabled = True
    Else
        BtnMnu(0).Enabled = False
        BtnMnu(2).Enabled = False
    End If
    
    If Trim(PayRcv.Text) <> "" And IsNumeric(PayRcv.Text) = True Then
        PayBal = Format(PayRcv - (DblServicePaid), "0.00")
    Else
        PayBal = "0.00"
    End If
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Service Controls Enable\Disable
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub SerControl(Enabled As Boolean)
    If Enabled = True Then
        SerBtn(1).Enabled = True
        SerQty.Enabled = True
    Else
        SerBtn(1).Enabled = False
        SerQty.Enabled = False
    End If
End Sub


'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' LogUsage Customer
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub LogUsage()
    Screen.MousePointer = 11
    
    SaveTransactionPos LvItems.ListItems
    DblServicePaid = 0
    
    Screen.MousePointer = 0
End Sub
