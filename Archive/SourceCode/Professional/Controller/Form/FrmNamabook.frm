VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmNamabook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nama Penunggu"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   30
      TabIndex        =   1
      Top             =   -75
      Width           =   3075
      Begin VB.ComboBox Combo3 
         Height          =   330
         ItemData        =   "FrmNamabook.frx":0000
         Left            =   2295
         List            =   "FrmNamabook.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1095
         Width           =   630
      End
      Begin VB.ComboBox Combo2 
         Height          =   330
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1095
         Width           =   720
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         ItemData        =   "FrmNamabook.frx":0016
         Left            =   750
         List            =   "FrmNamabook.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1095
         Width           =   615
      End
      Begin AIFCmp1.asxToolbar asxToolbar1 
         Height          =   420
         Left            =   2190
         Top             =   1515
         Width           =   780
         _ExtentX        =   1376
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
         ButtonCount     =   2
         CaptionOptions  =   0
         AutoSize        =   -1  'True
         ButtonKey1      =   "tambah"
         ButtonPicture1  =   "FrmNamabook.frx":0069
         ButtonToolTipText1=   "Tambah"
         ButtonKey2      =   "batal"
         ButtonPicture2  =   "FrmNamabook.frx":03BB
         ButtonToolTipText2=   "Batal"
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   105
         TabIndex        =   0
         ToolTipText     =   "Masukkan nama penunggu"
         Top             =   630
         Width           =   2820
      End
      Begin CafeBonzer.Label3D Label1 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Left            =   1410
         TabIndex        =   5
         Top             =   1125
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masa :"
         Height          =   210
         Left            =   135
         TabIndex        =   3
         Top             =   1140
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmNamabook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    tvindex = TvSelIndex(FrmBooking.Tv1)
    tvkey = TvSelKey(FrmBooking.Tv1)
     
    Select Case ButtonIndex
    Case 1
        If Text1 = "" Then Exit Sub
        If Combo1 = "" Then Exit Sub
        If Combo2 = "" Then Exit Sub
        If Combo3 = "" Then Exit Sub
        FrmBooking.Tv1.Nodes.Add tvkey, tvwChild, "time" & Combo1 & ":" & Combo2 & Combo3, UCase(Text1), "orang"
        Me.Hide
        FrmBooking.Enabled = True
        FrmBooking.SetFocus
        Unload Me
    Case 2
        Me.Hide
        FrmBooking.Enabled = True
        FrmBooking.SetFocus
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    tvindex = TvSelIndex(FrmBooking.Tv1)
    
    Label1.Caption = FrmBooking.Tv1.Nodes(tvindex).Text
    
    jam = Hour(Time)
    
    jamn = Right(Time, 2)
    
    For f = 0 To 59
    Combo2.AddItem f
    Next f
      
    Combo3 = "PM"
End Sub
