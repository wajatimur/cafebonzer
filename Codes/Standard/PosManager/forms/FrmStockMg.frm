VERSION 5.00
Begin VB.Form FrmStockMg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Management"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStockMg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox StockChk 
      Caption         =   "Enable stock management."
      Height          =   345
      Left            =   180
      TabIndex        =   1
      Top             =   30
      Value           =   1  'Checked
      Width           =   2730
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   3750
      Begin VB.TextBox InfoStockTxt 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   1
         Left            =   1890
         TabIndex        =   6
         Text            =   "0"
         Top             =   825
         Width           =   1635
      End
      Begin VB.TextBox InfoStockTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Index           =   0
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   1635
      End
      Begin CbSnmMgr.XpButton BtnStock 
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Add Stock"
         Top             =   1260
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
         MICON           =   "FrmStockMg.frx":23D2
         PICN            =   "FrmStockMg.frx":23EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CbSnmMgr.XpButton BtnStock 
         Height          =   375
         Index           =   1
         Left            =   2685
         TabIndex        =   9
         ToolTipText     =   "Deduct Stock"
         Top             =   1260
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
         MICON           =   "FrmStockMg.frx":2988
         PICN            =   "FrmStockMg.frx":29A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CbSnmMgr.XpButton BtnStock 
         Height          =   375
         Index           =   2
         Left            =   2250
         TabIndex        =   10
         ToolTipText     =   "Clear Stock"
         Top             =   1260
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
         MICON           =   "FrmStockMg.frx":2D3E
         PICN            =   "FrmStockMg.frx":2D5A
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
         Caption         =   "Stock :"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   405
         Width           =   1335
      End
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   450
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "Add Item"
      Top             =   1935
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
      MICON           =   "FrmStockMg.frx":32F4
      PICN            =   "FrmStockMg.frx":3310
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
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Add Item"
      Top             =   1935
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
      MICON           =   "FrmStockMg.frx":38AA
      PICN            =   "FrmStockMg.frx":38C6
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
Attribute VB_Name = "FrmStockMg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrItemID As String
Dim IntCurrentStock As Integer


Public Sub Init(StrItmID As String, StrCurStock As String)
    Load Me
    StrItemID = StrItmID
    If StrCurStock = Var(3) Then
        Call CControl(False)
        InfoStockTxt(0) = Var(3)
    Else
        IntCurrentStock = StrCurStock
        InfoStockTxt(0) = StrCurStock
    End If
    Me.Show vbModal
End Sub


Private Sub BtnMenu_Click(Index As Integer)
    Dim Rs As Recordset
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenDynaset)
    
    If Index = 1 Then GoTo Terminate
    
    With Rs
        .FindFirst "Id = '" & StrItemID & "'"
        If .NoMatch = False Then
            .Edit
            If StockChk.Value = 0 Then
                !Stock = -1
                FrmSnmMg.ItemInfoTxt(2) = Var(3)
            Else
                !Stock = InfoStockTxt(0)
                FrmSnmMg.ItemInfoTxt(2) = InfoStockTxt(0)
            End If
            .Update
        End If
    End With
    
Terminate:
    Set Rs = Nothing
    Unload Me
End Sub

Private Sub BtnStock_Click(Index As Integer)
    Select Case Index
    Case 0
        InfoStockTxt(0) = InfoStockTxt(1) + IntCurrentStock
    Case 1
        InfoStockTxt(0) = IntCurrentStock - InfoStockTxt(1)
    Case 2
        InfoStockTxt(0) = IntCurrentStock
    End Select
End Sub

Private Sub StockChk_Click()
    Call CControl(StockChk.Value)
End Sub

Public Sub CControl(Status As Boolean)
    If Status = True Then
        StockChk.Value = 1
    Else
        StockChk.Value = 0
    End If
    InfoStockTxt(1).Enabled = Status
    BtnStock(0).Enabled = Status
    BtnStock(1).Enabled = Status
    BtnStock(2).Enabled = Status
End Sub
