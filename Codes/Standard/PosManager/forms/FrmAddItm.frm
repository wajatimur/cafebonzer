VERSION 5.00
Begin VB.Form FrmItemAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Item"
   ClientHeight    =   2355
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   3885
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
   Icon            =   "FrmAddItm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkStock 
      Caption         =   "Stock :"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1380
      Width           =   1290
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   1860
      TabIndex        =   3
      Text            =   "0"
      Top             =   1335
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   915
      Width           =   1935
   End
   Begin VB.ComboBox CbGrp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      Height          =   330
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin CbSnmMgr.XpButton BtnMenu 
      Height          =   450
      Index           =   0
      Left            =   2895
      TabIndex        =   4
      ToolTipText     =   "Add Item"
      Top             =   1860
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
      MICON           =   "FrmAddItm.frx":23D2
      PICN            =   "FrmAddItm.frx":23EE
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
      Left            =   3375
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   1860
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
      MICON           =   "FrmAddItm.frx":2988
      PICN            =   "FrmAddItm.frx":29A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CbSnmMgr.Line3D Line3D1 
      Height          =   45
      Left            =   15
      TabIndex        =   9
      Top             =   1770
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Category :"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "FrmItemAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkStock_Click()
    Txt1(2).Enabled = ChkStock.Value
End Sub

Private Sub Form_Load()
    Dim Rs As Recordset, StrGrpId As String
    
    Set Rs = uSDB.OpenRecordset("ServiceCategory", dbOpenDynaset)
    If Rs.BOF = True Then Exit Sub
  '{ Loading POS category }
    With Rs
        .MoveFirst
        Do Until .EOF = True
            CbGrp.AddItem !Id & "-" & !Name
            .MoveNext
        Loop
    End With
    
  '{ Select a proper group }
    StrGrpId = FrmSnmMg.IcGroup.SelectedItem.Key
    For l = 0 To CbGrp.ListCount - 1
        If StrGrpId = Left(CbGrp.List(l), 5) Then
            CbGrp.ListIndex = l
            Exit Sub
        End If
    Next l
End Sub

Private Sub BtnMenu_Click(Index As Integer)
    Dim IntStock As Integer
    Select Case Index
        Case 0
            IntStock = Txt1(2)
            If CbGrp.Text = "" Then Exit Sub
            If Txt1(0) = "" Then Exit Sub
            If Txt1(1) = "" Then Exit Sub
            If ChkStock.Value = 0 Then IntStock = -1
            
            Call ItemAdd(Left(CbGrp, 5), Txt1(0), Txt1(1), IntStock)
            Txt1(0) = "": Txt1(1) = "": Txt1(2) = 0
        Case 1
            Set FrmAddItm = Nothing
            Unload Me
    End Select
End Sub
