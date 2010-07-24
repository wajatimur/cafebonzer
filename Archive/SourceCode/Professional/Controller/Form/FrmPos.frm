VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmPosMg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Services & Merchandise Manager"
   ClientHeight    =   5745
   ClientLeft      =   255
   ClientTop       =   1980
   ClientWidth     =   9735
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PosID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   6345
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   270
      Width           =   2940
   End
   Begin CafeBonzer.Line3D uLine3D2 
      Height          =   45
      Left            =   45
      TabIndex        =   7
      Top             =   -15
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CafeBonzer.Line3D uLine3D1 
      Height          =   45
      Left            =   4290
      TabIndex        =   6
      Top             =   4890
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin MSComctlLib.StatusBar PosStat 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   5430
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox PosHarga 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   6345
      TabIndex        =   4
      Top             =   1305
      Width           =   2940
   End
   Begin VB.TextBox PosItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   6345
      TabIndex        =   3
      Top             =   795
      Width           =   2940
   End
   Begin MSComctlLib.ImageList Iml 
      Left            =   3660
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":058A
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":0B26
            Key             =   "user"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":157A
            Key             =   "akses"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":1B16
            Key             =   "item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":20B2
            Key             =   "services"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":264E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":2766
            Key             =   "foods"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":2B02
            Key             =   "beverages"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":2E9E
            Key             =   "magazines"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":343A
            Key             =   "others"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPos.frx":39D6
            Key             =   "none"
         EndProperty
      EndProperty
   End
   Begin AIFCmp1.asxToolbar PosAsx1 
      Height          =   420
      Left            =   7725
      Top             =   4980
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   741
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   6
      CaptionOptions  =   0
      ButtonKey1      =   "tambah"
      ButtonPicture1  =   "FrmPos.frx":3F72
      ButtonToolTipText1=   "Add Item"
      ButtonKey2      =   "buang"
      ButtonPicture2  =   "FrmPos.frx":42C4
      ButtonToolTipText2=   "Delete Item"
      ButtonKey3      =   "Save"
      ButtonPicture3  =   "FrmPos.frx":4616
      ButtonToolTipText3=   "Save"
      ButtonStyle4    =   2
      ButtonKey5      =   "grpadd"
      ButtonPicture5  =   "FrmPos.frx":4968
      ButtonToolTipText5=   "Add Group"
      ButtonKey6      =   "grpdel"
      ButtonPicture6  =   "FrmPos.frx":4CBA
      ButtonToolTipText6=   "Delete Group"
   End
   Begin MSComctlLib.TreeView Tv1 
      Height          =   5355
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "Iml"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label PosLBL3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item ID :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4590
      TabIndex        =   9
      Top             =   330
      Width           =   810
   End
   Begin VB.Label PosLBL1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   4590
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label PosLBL2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Per Unit :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4590
      TabIndex        =   0
      Top             =   1380
      Width           =   1290
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Begin VB.Menu menu1exp 
         Caption         =   "Expand All"
      End
      Begin VB.Menu menu1col 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menu1close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Tools"
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Group Add"
         Index           =   0
      End
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Group Delete"
         Index           =   1
      End
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Item Add"
         Index           =   2
      End
      Begin VB.Menu menu2itemfunc 
         Caption         =   "Item Delete"
         Index           =   3
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2save 
         Caption         =   "Save Changes"
      End
   End
End
Attribute VB_Name = "FrmPosMg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs As Recordset


Private Sub Form_Load()
    Call LoadItemsGroups
End Sub


Private Sub menu1close_Click()
    Set FrmPosMg = Nothing
    Unload Me
End Sub


Private Sub menu1col_Click()
    For t = 1 To Tv1.Nodes.Count
        If Left(Tv1.Nodes(t).Key, 1) = "g" Then
            Tv1.Nodes(t).Expanded = False
        End If
    Next t
End Sub


Private Sub menu1exp_Click()
    For t = 1 To Tv1.Nodes.Count
        If Left(Tv1.Nodes(t).Key, 1) = "g" Then
            Tv1.Nodes(t).Expanded = True
        End If
    Next t
End Sub


Private Sub menu2itemfunc_Click(Index As Integer)
    Select Case Index
        Case 0
            Call PosAsx1_ButtonClick(5, "grpadd")
        Case 1
            Call PosAsx1_ButtonClick(6, "grpdel")
        Case 2
            Call PosAsx1_ButtonClick(1, "tambah")
        Case 3
            Call PosAsx1_ButtonClick(2, "buang")
    End Select
End Sub


Private Sub PosAsx1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Dim SelGrp As String, SelItm As Node
    
    Select Case ButtonIndex
    Case 1  'Add items
        FrmAddPos.Show vbModal
        LoadItemsGroups
        
    Case 2  'Delete items
        If SelItm Is Nothing Then GoTo ErrInt
        Set SelItm = Tv1.SelectedItem
        If SelItm Is Nothing Then Exit Sub
        If Left(SelItm.Key, 1) = "g" Or SelItm.Text = "" Then
            PosStat.Panels(2) = MB(15)
            Exit Sub
        End If
        lret = MsgBox(MB(2) & " " & SelItm.Text & " ?", vbOKCancel, CbMsgWarn)
        If lret = vbCancel Then Exit Sub
        Tv1.Nodes.Remove (SelItm.Index)
        uSDBe.DataRemove "pos-items", "nama", SelItm.Text
        
    
    Case 3  'Save items changes
        Set SelItm = Tv1.SelectedItem
        If SelItm Is Nothing Then GoTo ErrInt
        If Left(SelItm.Key, 1) = "g" Or SelItm.Text = "" Then
            PosStat.Panels(2) = MB(15)
            Exit Sub
        ElseIf PosItem = "" Or PosHarga = "" Then
            PosStat.Panels(2) = VS(10)
            Exit Sub
        End If
        
        uSDBe.DataEdit "pos-items", "nama", "id", SelItm.Key, PosItem, True, False
        uSDBe.DataEdit "pos-items", "harga", "id", SelItm.Key, PosHarga, False, True
        LoadItemsGroups
        
    Case 5  'Add a group
        s_grp$ = MgoInpt.GetInput("Group Name", BtnClose)
        If Trim(s_grp$) <> "" Then posAddGroup (s_grp$)
        LoadItemsGroups
        
    Case 6  'Delete a group
        If Tv1.SelectedItem Is Nothing Then GoTo ErrInt
        SelGrp = Tv1.SelectedItem.Key

        If Left(SelGrp, 1) = "g" Then
            SelGrp = Mid(SelGrp, 2)
            ret = MsgBox(MB(18), vbOKCancel, CbMsgWarn)
            If ret = vbOK Then
                ret = posDelGroup(SelGrp)
                If ret > 0 Then Tv1.Nodes.Remove ("g" & SelGrp)
            End If
        Else
            PosStat.Panels(2) = MB(19)
        End If
    End Select
    
    PosStat.Panels(1) = ""
    PosStat.Panels(2) = ""
    PosID = "": PosItem = "": PosHarga = ""
Exit Sub

ErrInt:
    If ButtonIndex = 3 Or ButtonIndex = 2 Then PosStat.Panels(2) = "Please select an item !"
    If ButtonIndex = 6 Then PosStat.Panels(2) = "Please select a group !"
End Sub


Private Sub Tv1_NodeClick(ByVal Node As MSComctlLib.Node)
    PosStat.Panels(2) = ""
    PosID = ""
    PosItem = ""
    PosHarga = ""
    If Left(Node.Key, 1) <> "g" Then
        PosID = Node.Key
        PosItem = Node.Text
        PosHarga = Node.Tag
    End If
End Sub


Private Sub LoadItemsGroups()
On Error GoTo ErrInt
    Dim ErrPos As Long
    Dim IconName As String
    
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenSnapshot)
    Rs.Sort = "id"
    Set Rs = Rs.OpenRecordset
    Tv1.Nodes.Clear
    
    'checking group counts
    If Rs.BOF = True Then
        MsgBox MB(16), vbInformation, CbMsgWarn
        menu1exp.Enabled = False
        menu1col.Enabled = False
        menu2itemadd.Enabled = False
        menu2itemdel.Enabled = False
        menu2grpdel.Enabled = False
        menu2save.Enabled = False
        Exit Sub
    End If
    
    'loading groups
    ErrPos = 1
    With Rs
        .MoveFirst
        Do Until .EOF = True
            IconName = "folder"
            
            For t = 1 To Iml.ListImages.Count
                If LCase(Iml.ListImages(t).Key) = LCase(!Name) Then IconName = LCase(!Name)
            Next t
            Tv1.Nodes.Add , , "g" & !id, !Name, IconName, IconName
            .MoveNext
        Loop
    End With
    
    'loading items
    ErrPos = 2
    For d = 0 To uSDBe.DataCount("pos-items") - 1
ErrPos2:
        grp = uSDBe.DataGet("pos-items", "groupid", d)
        id = uSDBe.DataGet("pos-items", "id", d)
        nme = uSDBe.DataGet("pos-items", "nama", d)
        hrg = uSDBe.DataGet("pos-items", "harga", d)
        cc = Tv1.Nodes("g" & grp).Children + 1
        Tv1.Nodes.Add "g" & grp, tvwChild, id, nme, "item"  '"item" & grp & cc, nme, "item"
        Tv1.Nodes(id).Tag = hrg
    Next d
    
    PosStat.Panels(1) = d & " items loaded !"
Exit Sub

ErrInt:
    If errnum = 35602 Then
        If ErrPos = 2 Then d = d + 1: GoTo ErrPos2
    End If
End Sub


Private Function ItemIdCheck(GroupId) As String

End Function
