VERSION 5.00
Begin VB.Form FrmAddItm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1350
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4290
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   915
      Width           =   1935
   End
   Begin VB.ComboBox CbGrp 
      Height          =   330
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox GuiBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   3915
      ScaleHeight     =   1755
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   0
      Width           =   375
      Begin CbSnmMgr.XpButton BtnMenu 
         Height          =   345
         Index           =   0
         Left            =   15
         TabIndex        =   7
         ToolTipText     =   "Add Item"
         Top             =   630
         Width           =   360
         _ExtentX        =   635
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmAddItm.frx":0000
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
         Height          =   345
         Index           =   1
         Left            =   15
         TabIndex        =   8
         ToolTipText     =   "Add Item"
         Top             =   990
         Width           =   360
         _ExtentX        =   635
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmAddItm.frx":001C
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   525
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Category :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAddItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim Rs As Recordset, strGrpId As String
    
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenDynaset)
    If Rs.BOF = True Then Exit Sub
  ' loading POS category
    With Rs
        .MoveFirst
        Do Until .EOF = True
            CbGrp.AddItem !id & "-" & !Name
            .MoveNext
        Loop
    End With
    
  ' select a proper group
    strGrpId = Mid(FrmSnmMg.IcGroup.SelectedItem.Key, 2)
    For l = 0 To CbGrp.ListCount - 1
        If strGrpId = Left(CbGrp.List(l), 2) Then
            CbGrp.ListIndex = l
            Exit Sub
        End If
    Next l
End Sub

Private Sub BtnMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            If CbGrp.Text = "" Then Exit Sub
            If Txt1(0) = "" Then Exit Sub
            If Txt1(1) = "" Then Exit Sub
        
            Call ItemAdd(Left(CbGrp, 2), Txt1(0), Txt1(1))
            Txt1(0) = "": Txt1(1) = ""
        Case 1
            Set FrmAddItm = Nothing
            Unload Me
    End Select
End Sub

Private Sub GuiBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveForm Me.hwnd
End Sub
