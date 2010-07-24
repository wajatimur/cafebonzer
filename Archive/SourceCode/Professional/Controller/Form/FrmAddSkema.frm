VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmAddSkema 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
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
   Icon            =   "FrmAddSkema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txt1 
      Height          =   315
      Index           =   1
      Left            =   1650
      TabIndex        =   1
      Top             =   2940
      Width           =   1770
   End
   Begin VB.TextBox Txt1 
      Height          =   315
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   2565
      Width           =   1770
   End
   Begin MSComctlLib.ListView Lv1 
      Height          =   2385
      Left            =   75
      TabIndex        =   3
      Top             =   90
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   4207
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Skema harga"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Harga per minit"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      Height          =   3375
      Left            =   3525
      ScaleHeight     =   3375
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   -15
      Width           =   360
      Begin AIFCmp1.asxToolbar Asx1 
         Height          =   390
         Left            =   0
         Top             =   2955
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   1
         ButtonKey1      =   "ok"
         ButtonPicture1  =   "FrmAddSkema.frx":000C
         ButtonToolTipText1=   "Ok/Cancel"
      End
      Begin AIFCmp1.asxToolbar Asx2 
         Height          =   390
         Left            =   0
         Top             =   2265
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   1
         ButtonKey1      =   "tambah"
         ButtonPicture1  =   "FrmAddSkema.frx":035E
         ButtonToolTipText1=   "Add Scheme"
      End
      Begin AIFCmp1.asxToolbar asxToolbar1 
         Height          =   390
         Left            =   0
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   1
         ButtonKey1      =   "padam"
         ButtonPicture1  =   "FrmAddSkema.frx":06B0
         ButtonToolTipText1=   "Delete Scheme"
      End
      Begin AIFCmp1.asxToolbar AsxC 
         Height          =   390
         Left            =   0
         Top             =   2625
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   1
         ButtonKey1      =   "batal"
         ButtonPicture1  =   "FrmAddSkema.frx":0A02
         ButtonToolTipText1=   "Cancel"
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Price per Minute :"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   2970
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scheme Name :"
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   2595
      Width           =   1395
   End
End
Attribute VB_Name = "FrmAddSkema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SADB As New clsData

Private Sub Asx1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Unload FrmAddSkema
    Set SADB = Nothing
    Set FrmAddSkema = Nothing
End Sub

Private Sub asx2_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Dim lItm As ListItem
    Dim itmfind As ListItem
    If Txt1(0) = "" Then Exit Sub
    If Txt1(1) = "" Then Exit Sub
    
    Set itmfind = Lv1.FindItem(Txt1(0))
    If itmfind Is Nothing Then
        Set lItm = Lv1.ListItems.Add(, , Txt1(0))
        lItm.SubItems(1) = Txt1(1)
        'tambah dalam database
        SADB.DataSave "skema", "skema", Txt1(0), True, False
        SADB.DataSave "Skema", "harga", Txt1(1), False, True
    Else
        MsgBox MB(1), vbOKOnly, CbMsgWarn: Exit Sub
    End If

    Txt1(0) = ""
    Txt1(1) = ""
End Sub

Private Sub AsxC_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Unload Me
End Sub

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    If Lv1.ListItems.Count = 0 Then Exit Sub
    If Lv1.SelectedItem.Text = "" Then Exit Sub
    
    nm = Lv1.SelectedItem.Text
    lret = MsgBox(MB(2) & " " & nm & " ?", vbOKCancel, CbMsgWarn)
    If lret = vbCancel Then Exit Sub
    Lv1.ListItems.Remove (Lv1.SelectedItem.Index)
    SADB.DataRemove "skema", "skema", nm
End Sub

Private Sub Form_Load()
    Dim lItm As ListItem
    
    SADB.InitDb = App.Path & "\data\data.mdb"
    For n = 0 To SADB.DataCount("skema") - 1
        sk = SADB.DataGet("skema", "skema", n)
        hg = SADB.DataGet("skema", "harga", n)
        Set lItm = Lv1.ListItems.Add(, , sk)
        lItm.SubItems(1) = hg
    Next n
    
    Set lItm = Nothing
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveFrm FrmAddSkema.hwnd
End Sub

