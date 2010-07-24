VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmHarga 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5340
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6675
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin CafeBonzer.Line3D uLine3D2 
      Height          =   45
      Left            =   3570
      TabIndex        =   23
      Top             =   3435
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.TextBox Baki 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Endless Showroom"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   510
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Just press Enter if the value same as above."
      Top             =   2790
      Width           =   2385
   End
   Begin VB.Frame Frame2 
      Caption         =   "Services && Merchandise"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   75
      TabIndex        =   13
      Top             =   1470
      Width           =   3450
      Begin VB.VScrollBar SerScroll1 
         Height          =   330
         Left            =   2085
         Max             =   999
         Min             =   1
         TabIndex        =   24
         Top             =   1290
         Value           =   999
         Width           =   165
      End
      Begin VB.TextBox JumlahSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3450
         Width           =   2085
      End
      Begin MSComctlLib.ListView LvSer 
         Height          =   1530
         Left            =   105
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1845
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   2699
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
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
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty."
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.TextBox QtySer 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   "1"
         Top             =   1305
         Width           =   780
      End
      Begin CafeBonzer.Line3D uLine3D1 
         Height          =   45
         Left            =   105
         TabIndex        =   18
         Top             =   1695
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin MSComctlLib.ImageCombo ImgCmb1 
         Height          =   330
         Left            =   1275
         TabIndex        =   1
         Top             =   330
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   16761024
         Locked          =   -1  'True
         Text            =   "None"
      End
      Begin MSComctlLib.ImageCombo ImgCmb2 
         Height          =   330
         Left            =   1275
         TabIndex        =   2
         Top             =   780
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   16761024
         Locked          =   -1  'True
         Text            =   "None"
      End
      Begin CafeBonzer.XpButton SerBtn 
         Height          =   360
         Index           =   0
         Left            =   2385
         TabIndex        =   25
         Top             =   1275
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
         MICON           =   "FrmHarga.frx":000C
         PICN            =   "FrmHarga.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SerBtn 
         Height          =   360
         Index           =   1
         Left            =   2820
         TabIndex        =   26
         Top             =   1275
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
         MICON           =   "FrmHarga.frx":05C2
         PICN            =   "FrmHarga.frx":05DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   570
         TabIndex        =   21
         Top             =   3480
         Width           =   645
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   150
         TabIndex        =   19
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Items :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   195
         TabIndex        =   15
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Category :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   195
         TabIndex        =   14
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.TextBox Bayar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Endless Showroom"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   4050
      TabIndex        =   4
      ToolTipText     =   "Just press Enter if the value same as above."
      Top             =   3900
      Width           =   2385
   End
   Begin VB.TextBox Jumlah 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Endless Showroom"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   510
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "RM 0.00"
      Top             =   1875
      Width           =   2400
   End
   Begin VB.Frame Frame1 
      Caption         =   "PC Usage"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   6510
      Begin VB.TextBox Harga 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1230
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   840
         Width           =   1860
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Height          =   360
         Left            =   1230
         ScaleHeight     =   330
         ScaleWidth      =   1815
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
         Width           =   1845
         Begin CafeBonzer.Label3D Masa 
            Height          =   210
            Left            =   45
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   45
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   370
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor2      =   16761024
            Caption         =   "10 Jam 10 Minit"
            Alignment       =   2
            BackColor       =   8421504
         End
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   90
         Picture         =   "FrmHarga.frx":0B78
         Top             =   855
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   105
         Picture         =   "FrmHarga.frx":1102
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Price :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Top             =   870
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   420
         TabIndex        =   10
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.Image MainBtn 
      Height          =   480
      Index           =   1
      Left            =   6210
      Picture         =   "FrmHarga.frx":168C
      ToolTipText     =   "Accept"
      Top             =   4875
      Width           =   480
   End
   Begin VB.Image MainBtn 
      Height          =   480
      Index           =   0
      Left            =   5820
      Picture         =   "FrmHarga.frx":36FE
      ToolTipText     =   "Cancel"
      Top             =   4875
      Width           =   480
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   3720
      TabIndex        =   17
      Top             =   2445
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   3720
      TabIndex        =   12
      Top             =   3555
      Width           =   1035
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   3705
      TabIndex        =   9
      Top             =   1590
      Width           =   810
   End
End
Attribute VB_Name = "FrmHarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs As Recordset

Public PrePaid As Boolean
Public PcName As String
Public pcCusName As String
Public pcInTime As String
Public pcOutTime As String
Public pcTotalTime As String
Public pcPaid As String
Public SerTotal As Double


Private Sub Bayar_Change()
    If Trim(Bayar.Text) <> "" And IsNumeric(Bayar.Text) = True Then
        Baki = Format(Bayar - (pcPaid + SerTotal), "#0.00")
    Else
        Baki = ""
    End If
End Sub

Private Sub Bayar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Log
        Unload FrmHarga
    End If
End Sub


Private Sub Form_Activate()
    Harga.SetFocus
    Harga.SelStart = 0
    Harga.SelLength = Len(Harga)
End Sub


Private Sub Form_Load()
    Dim ImL1 As ImageList
    Dim iconName As String
    
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenSnapshot)
    Set ImL1 = FrmMain.Iml
    ImgCmb1.ImageList = ImL1
    ImgCmb1.ComboItems.Clear
    
    If SetAmbil("tukarharga") = 0 Then Harga.Locked = True
    If Rs.BOF = True Then Exit Sub
    
    With Rs
        .MoveFirst
        Do Until .EOF = True
            iconName = "folder"
            
            For t = 1 To ImL1.ListImages.Count
                If LCase(ImL1.ListImages(t).Key) = LCase(!Name) Then iconName = LCase(!Name)
            Next t
            ImgCmb1.ComboItems.Add , "g" & !id, !Name, iconName, iconName
            .MoveNext
        Loop
    End With
    
    ImgCmb1.ComboItems.Add , VS(1), VS(1), "none", "none"
    ImgCmb1.ComboItems.Item(VS(1)).Selected = True
End Sub


Private Sub Harga_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Harga = "" Then
            ret = MsgBox("Void this PC usage ?", vbOKCancel, CbMsgWarn)
            If ret = vbOK Then
                Unload FrmHarga
            Else
                Harga = pcPaid
                Harga.SelStart = 1
                Harga.SelLength = Len(Harga)
                Exit Sub
            End If
        End If
        If IsNumeric(Harga) = False Then
            MsgBox MB(4), vbInformation, CbMsgWarn
            Harga.SetFocus
            Harga = pcPaid
            Exit Sub
        End If
        
        pcPaid = Harga
        Jumlah = Crnc & " " & Format(pcPaid + SerTotal, "#0.00")
        Bayar.SetFocus
    End If
End Sub

Private Sub Harga_LostFocus()
    Call Harga_KeyUp(13, 1)
End Sub


Private Sub ImgCmb1_Change()
    Dim Rss As Recordset
    Dim CbItm As ComboItem
    Set Rss = uSDB.OpenRecordset("pos-items", dbOpenSnapshot)
    
    If ImgCmb1.Text <> VS(1) Then
        ImgCmb2.Enabled = True
        QtySer.Enabled = True
        SerBtn(0).Enabled = True
        SerBtn(1).Enabled = True
        LvSer.Enabled = True
        SerScroll1.Enabled = True
        GoTo LoadItem
    Else
        ImgCmb2.Text = VS(1)
        ImgCmb2.Enabled = False
        QtySer.Enabled = False
        SerBtn(0).Enabled = False
        SerBtn(1).Enabled = False
        LvSer.Enabled = False
        SerScroll1.Enabled = False
        Exit Sub
    End If
    
LoadItem:
    ImgCmb2.ComboItems.Clear
    Rss.Filter = "groupid = '" & Mid(ImgCmb1.SelectedItem.Key, 2) & "'"
    
    Set Rs = Rss.OpenRecordset
    
    If Rs.BOF = True Then Exit Sub
    With Rs
        .MoveFirst
        Do Until .EOF = True
            Set CbItm = ImgCmb2.ComboItems.Add(, !id, !Nama)
            CbItm.Tag = !Harga
            .MoveNext
        Loop
    End With
    
    'resize the list width
    For g = 1 To ImgCmb2.ComboItems.Count
        tWd = TextWidth(ImgCmb2.ComboItems(g).Text)
        If tWd > tInt Then tInt = tWd
    Next g
    tInt = (tInt / 15) + 40
    ret = SendMessage(ImgCmb2.hwnd, CB_SETDROPPEDWIDTH, tInt, 0)
    
    'default selection
    ImgCmb2.ComboItems(1).Selected = True
End Sub

Private Sub ImgCmb1_Click()
    Call ImgCmb1_Change
End Sub


Private Sub LvSer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ItemTotal As Double
    
    ItemTotal = CDbl(Item.SubItems(1)) * CInt(Item.SubItems(2))
    JumlahSer = Crnc & " " & ItemTotal & " / " & Crnc & " " & SerTotal
End Sub


Private Sub Log()
    Dim dDay
    
    dDay = Weekday(Date)
    dDay = Choose(dDay, "Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu")
    Screen.MousePointer = 11

 '{ save dalam table usage }'
    SavePcUsage PcName, pcCusName, pcInTime, pcOutTime, pcPaid
 '{ save dalam table bulanan }'
    SavePcBulanan PcName, pcTotalTime, pcPaid
 '{ rekod jualan harian }'
    SavePcHarian pcPaid
 '{ save dalam table graf-mingguan(untuk mengira hari graf hari) }'
    SavePcMingguan dDay, pcPaid
 '{ save dalam table pelanggan (!!!tolong pikir skit untuk CusID tuh) }'
    SavePelanggan pcCusName, "", pcTotalTime, pcPaid
 '{ save transaksi POS }'
    SavePosTrans LvSer.ListItems
    
    'reset agent, main ui..
    SelAgn.CusStop
    SelAgn.NetSend "//logout"
    
    SerTotal = 0
    pcPaid = 0
    pcInTime = ""
    pcOutTime = ""
    pcTotalTime = 0
    PcName = ""
    pcCusName = ""
    Screen.MousePointer = 0
End Sub

Private Sub MainBtn_Click(Index As Integer)
    Dim cItm As ListItem
    Set cItm = FrmMain.Lv1.SelectedItem
    
    Select Case Index
    Case 0
        'cancel
        cItm.SubItems(3) = VS(1)
    Case 1
        Call Harga_KeyUp(13, 1)
        Call Log
        For t = 1 To FrmMain.Lv1.ListItems.Count
            RecoveryGo FrmMain.Lv1.ListItems(t)
        Next t
        Call UpdatePanel(SelText)
    End Select
    
    FrmHost.Timer1.Enabled = True
    Unload FrmHarga
End Sub

Private Sub SerBtn_Click(Index As Integer)
    Dim SCbItm As ComboItem, lvItm As ListItem, fItm As ListItem
     
    Select Case Index
    Case 0
        Set SCbItm = ImgCmb2.SelectedItem
        If QtySer > 0 Then
            Set fItm = LvSer.FindItem(SCbItm.Text)
            
            If fItm Is Nothing Then
                Set lvItm = LvSer.ListItems.Add(, SCbItm.Key, SCbItm.Text)
                lvItm.SubItems(1) = SCbItm.Tag
                lvItm.SubItems(2) = QtySer
            Else
                fItm.SubItems(2) = QtySer
            End If
        Else
            MsgBox "Please enter quantity !", vbOKOnly, CbMsgWarn
            QtySer.SelStart = 1
            QtySer.SelLength = Len(QtySer.Text)
            Exit Sub
        End If
    Case 1
        If LvSer.ListItems.Count = 0 Then Exit Sub
        LvSer.ListItems.Remove (LvSer.SelectedItem.Index)
    End Select
    
    'recalculate total
    SerTotal = 0
    For g = 1 To LvSer.ListItems.Count
        SerTotal = SerTotal + (CDbl(LvSer.ListItems(g).SubItems(1)) * CInt(LvSer.ListItems(g).SubItems(2)))
    Next g
    JumlahSer = Crnc & " " & Format$(SerTotal, "#0.00")
    
    'recalculate overall total
     Jumlah = Crnc & " " & Format(pcPaid + SerTotal, "#0.00")
End Sub

Private Sub SerScroll1_Change()
    QtySer = 1000 - SerScroll1.Value
End Sub
