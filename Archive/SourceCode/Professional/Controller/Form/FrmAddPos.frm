VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmAddPos 
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
   Icon            =   "FrmAddPos.frx":0000
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
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "FrmAddPos.frx":000C
      Left            =   1860
      List            =   "FrmAddPos.frx":000E
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
   Begin VB.PictureBox Picture5 
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
      Left            =   3930
      ScaleHeight     =   1755
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   0
      Width           =   360
      Begin AIFCmp1.asxToolbar Asx1 
         Height          =   390
         Left            =   0
         Top             =   975
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
         ButtonPicture1  =   "FrmAddPos.frx":0010
         ButtonToolTipText1=   "Ok/Cancel"
      End
      Begin AIFCmp1.asxToolbar AsxC 
         Height          =   390
         Left            =   0
         Top             =   645
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
         ButtonPicture1  =   "FrmAddPos.frx":0362
         ButtonToolTipText1=   "Cancel"
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
      Caption         =   "Item :"
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
Attribute VB_Name = "FrmAddPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Dim nNod As Node, lCnt As Long
    
    If Combo1.Text = "" Then Exit Sub
    If Txt1(0) = "" Then Exit Sub
    If Txt1(1) = "" Then
        Exit Sub
    Else
        gidx = Format(Left(Combo1.Text, 2), "#00")
        'iidx = FrmPosMg.Tv1.Nodes("g" & gidx).Children + 1
        iidx = lCnt + 1
        iidx = Format(iidx, "#000")
        'For a% = 1 To FrmPosMg.Tv1.Nodes("g" & gidx).Children
        
            
        Set nNod = FrmPosMg.Tv1.Nodes.Add("g" & gidx, tvwChild, , Txt1(0), "item")
        nNod.Tag = Txt1(1)
        uSDBe.DataSave "pos-items", "groupid", gidx, True, False
        uSDBe.DataSave "pos-items", "id", "p" & gidx & iidx, False, False
        uSDBe.DataSave "pos-items", "nama", Txt1(0), False, False
        uSDBe.DataSave "pos-items", "harga", Txt1(1), False, True
        Txt1(0) = "": Txt1(1) = ""
    End If
End Sub

Private Sub AsxC_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Unload Me
End Sub


Private Sub Form_Load()
    'loading POS category
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenDynaset)
    
    If Rs.BOF = True Then Exit Sub
    With Rs
        .MoveFirst
        Do Until .EOF = True
            Combo1.AddItem !id & !Name
            .MoveNext
        Loop
    End With
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveFrm Me.hwnd
End Sub
