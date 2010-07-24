VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmLogout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5280
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4845
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
   Icon            =   "FrmHarga.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1260
      TabIndex        =   0
      ToolTipText     =   "Enter Received Value."
      Top             =   4725
      Width           =   1515
   End
   Begin VB.PictureBox PayBox 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   2700
      ScaleHeight     =   735
      ScaleWidth      =   2145
      TabIndex        =   15
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   45
         Width           =   675
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3780
      Left            =   60
      TabIndex        =   7
      Top             =   810
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   6668
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Usage"
      TabPicture(0)   =   "FrmHarga.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "BtnItems(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LvItems"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Add On"
      TabPicture(1)   =   "FrmHarga.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SerScroll1"
      Tab(1).Control(1)=   "SerQty"
      Tab(1).Control(2)=   "SerICmb"
      Tab(1).Control(3)=   "SerLv"
      Tab(1).Control(4)=   "SerBtn(1)"
      Tab(1).Control(5)=   "SerLbl(0)"
      Tab(1).Control(6)=   "SerLbl(2)"
      Tab(1).ControlCount=   7
      Begin MSComctlLib.ListView LvItems 
         Height          =   2820
         Left            =   105
         TabIndex        =   20
         Top             =   435
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   4974
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Desription"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Total"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.VScrollBar SerScroll1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -71010
         Max             =   999
         Min             =   1
         TabIndex        =   12
         Top             =   3330
         Value           =   999
         Width           =   165
      End
      Begin VB.TextBox SerQty 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71835
         TabIndex        =   11
         Text            =   "1"
         Top             =   3345
         Width           =   780
      End
      Begin MSComctlLib.ImageCombo SerICmb 
         Height          =   330
         Left            =   -73800
         TabIndex        =   8
         Top             =   450
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Locked          =   -1  'True
         Text            =   "None"
      End
      Begin MSComctlLib.ListView SerLv 
         Height          =   2370
         Left            =   -74910
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   900
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   4180
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   -70800
         TabIndex        =   10
         Top             =   3330
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
         MICON           =   "FrmHarga.frx":0940
         PICN            =   "FrmHarga.frx":095C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton BtnItems 
         Height          =   345
         Index           =   0
         Left            =   4215
         TabIndex        =   22
         Top             =   3330
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
         MICON           =   "FrmHarga.frx":0EF6
         PICN            =   "FrmHarga.frx":0F12
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
         BackStyle       =   0  'Transparent
         Caption         =   "Category :"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label SerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   -72870
         TabIndex        =   13
         Top             =   3375
         Width           =   960
      End
   End
   Begin VB.PictureBox LgBanner 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   2700
      TabIndex        =   4
      Top             =   0
      Width           =   2700
      Begin CafeBonzer.Label3D LgBannerTerminal 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "Nama Pc"
         Top             =   390
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   12632256
         ForeColor2      =   0
         Caption         =   "Agent"
         BackColor       =   16777215
      End
      Begin VB.Image LgBannerImg 
         Height          =   960
         Left            =   75
         Picture         =   "FrmHarga.frx":14AC
         Top             =   30
         Width           =   960
      End
      Begin VB.Label LgBannerLogout 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logout"
         Height          =   195
         Left            =   2070
         TabIndex        =   6
         Top             =   30
         Width           =   570
      End
   End
   Begin CafeBonzer.Line3D Line3D1 
      Height          =   45
      Index           =   0
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CafeBonzer.XpButton BtnMnu 
      Height          =   585
      Index           =   0
      Left            =   4155
      TabIndex        =   1
      ToolTipText     =   "Confirm Transaction"
      Top             =   4650
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
      MICON           =   "FrmHarga.frx":1F45
      PICN            =   "FrmHarga.frx":1F61
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
      Left            =   3465
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   4650
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
      MICON           =   "FrmHarga.frx":24FB
      PICN            =   "FrmHarga.frx":2517
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   105
      TabIndex        =   21
      Top             =   4725
      Width           =   1005
   End
End
Attribute VB_Name = "FrmLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmLogout
'    Project    : CafeBonzer
'
'    Description: Customer Logout
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private StrTerminal As String       'LoadUsage, LogUsage
Private StrCustomer As String       'LoadUsage, LogUsage
Private StrCustomerId As String     'LoadUsage, LogUsage
Private StrTimeIn As String         'LoadUsage, LogUsage
Private StrTimeOut As String        'LoadUsage, LogUsage
Private LngUsageTimeMinute As Long  'LoadUsage, LogUsage

Private DblUsagePrice As Double     'CalculateTotal, LoadUsage, LogUsage
Private DblServicePaid As Double    'CalculateTotal


Private Sub Form_Load()
    LvItems.SmallIcons = FrmMain.ImgList16
    
    'If SecAccessRequest(TaskChangePrice) = False Then PayPrice.Locked = True
    Call LoadPosCatCB(SerICmb, FrmMain.ImgListSnm, True)
    Call LoadUsage
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
        Call UpdatePanel(SelText)
        AgentSel.AgnRecoverRemove
    End Select
    Unload Me
End Sub

Private Sub SerScroll1_Change()
    SerQty = 1000 - SerScroll1.Value
End Sub

Private Sub SerBtn_Click(Index As Integer)
    Dim TmpItem As ListItem, SelItem As ListItem, FindItm As ListItem
    Set TmpItem = SerLv.SelectedItem
    
    If SerQty > 0 Then
        Set FindItm = LvItems.FindItem(TmpItem.Text)
        If FindItm Is Nothing Then
            Set SelItem = LvItems.ListItems.Add(, TmpItem.Key, TmpItem.Text, , "ITEM")
            SelItem.SubItems(1) = VS(2, 2) & " = " & SerQty
            SelItem.SubItems(2) = Format(SerQty * TmpItem.SubItems(1), "0.00")
            SelItem.Tag = TmpItem.Tag
        Else
            FindItm.SubItems(1) = VS(2, 2) & " = " & SerQty
            FindItm.SubItems(2) = Format(SerQty * TmpItem.SubItems(1), "0.00")
        End If
        Call CalculateTotal
    Else
        MsgBox ST(0, 0), vbOKOnly, CbMsgWarn
        SerQty.SelStart = 1
        SerQty.SelLength = Len(SerQty.Text)
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
    Dim LngItemCnt As Long
    
    DblServicePaid = 0
    LngItemCnt = LvItems.ListItems.Count
    For g = 1 To LngItemCnt
        If LvItems.ListItems(g).Text <> VS(2, 0) Then
            DblServicePaid = DblServicePaid + CDbl(LvItems.ListItems(g).SubItems(2))
        End If
    Next
    
    PayTotal = Format$(DblUsagePrice + DblServicePaid, "0.00")
    
    If Trim(PayRcv.Text) <> "" And IsNumeric(PayRcv.Text) = True Then
        PayBal = Format(PayRcv - (DblUsagePrice + DblServicePaid), "0.00")
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
        SerScroll1.Enabled = True
    Else
        SerBtn(1).Enabled = False
        SerQty.Enabled = False
        SerScroll1.Enabled = False
    End If
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Load Customer Usage
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub LoadUsage()
    Dim TmpItm As ListItem
    
    DblUsagePrice = AgentSel.CusGetUsage
    LngUsageTimeMinute = AgentSel.CusGetUsageTime(True)
    
    StrTimeIn = AgentSel.CustomerTimeIn
    If AgentSel.CustomerTimeOut = VS(0, 1) Then
        StrTimeOut = Now
    Else
        StrTimeOut = AgentSel.CustomerTimeOut
    End If
    
    StrCustomer = AgentSel.CustomerName
    StrCustomerId = AgentSel.CustomerId
    StrTerminal = AgentSel.AgentName
    
    Set TmpItm = LvItems.ListItems.Add(, , VS(2, 0), , "TIME")
    TmpItm.SubItems(1) = AgentSel.CusGetUsageTime
    TmpItm.SubItems(2) = Format(DblUsagePrice, "0.00")
    
    LgBannerTerminal.Caption = StrCustomer
    LgBannerLogout = SL(2) & " - " & StrTerminal
    Call CalculateTotal
End Sub

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' LogUsage Customer
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub LogUsage()
    Screen.MousePointer = 11
    
 '{ save dalam table pelanggan }'
    SaveCustomer StrCustomer, StrCustomerId, LngUsageTimeMinute, DblUsagePrice
 '{ save dalam table usage }'
    SaveTransactionPc StrTerminal, StrCustomer, StrTimeIn, StrTimeOut, DblUsagePrice
 '{ save transaksi POS }'
    SaveTransactionPos LvItems.ListItems

    AgentSel.CusStop
    DblUsagePrice = 0
    DblServicePaid = 0
    Screen.MousePointer = 0
End Sub
