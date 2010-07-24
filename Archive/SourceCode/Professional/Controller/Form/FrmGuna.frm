VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmGuna 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4320
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
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
      Height          =   3420
      Left            =   3960
      ScaleHeight     =   3420
      ScaleWidth      =   360
      TabIndex        =   14
      Top             =   0
      Width           =   360
      Begin CafeBonzer.XpButton MainBtnOk 
         Height          =   345
         Left            =   0
         TabIndex        =   17
         Top             =   3075
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   609
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
         MICON           =   "FrmGuna.frx":0000
         PICN            =   "FrmGuna.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton MainBtnKo 
         Height          =   345
         Left            =   0
         TabIndex        =   18
         Top             =   2745
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   609
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
         MICON           =   "FrmGuna.frx":05B6
         PICN            =   "FrmGuna.frx":05D2
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   30
      TabIndex        =   8
      Top             =   -75
      Width           =   3900
      Begin CafeBonzer.Line3D uLine3D1 
         Height          =   45
         Left            =   75
         TabIndex        =   16
         Top             =   1500
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin VB.ComboBox cbPaid 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
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
         ItemData        =   "FrmGuna.frx":0B6C
         Left            =   1125
         List            =   "FrmGuna.frx":0B7F
         TabIndex        =   4
         Text            =   "1.00"
         Top             =   2250
         Width           =   1080
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fixed Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Index           =   3
         Left            =   285
         TabIndex        =   5
         Top             =   2655
         Width           =   1710
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pre Paid"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Index           =   2
         Left            =   285
         TabIndex        =   3
         Top             =   1935
         Width           =   1710
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pay As U Go"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Index           =   1
         Left            =   285
         TabIndex        =   2
         Top             =   1605
         Value           =   -1  'True
         Width           =   1710
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmGuna.frx":0BA1
         Left            =   1905
         List            =   "FrmGuna.frx":0BA3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1035
         Width           =   1785
      End
      Begin CafeBonzer.Label3D NamaPc 
         Height          =   300
         Left            =   195
         TabIndex        =   12
         ToolTipText     =   "Nama Pc"
         Top             =   210
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   16777215
         ForeColor2      =   0
         Caption         =   "Agent"
         BackColor       =   12632256
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
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
         ItemData        =   "FrmGuna.frx":0BA5
         Left            =   1995
         List            =   "FrmGuna.frx":0BBB
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3015
         Width           =   840
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
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
         ItemData        =   "FrmGuna.frx":0BD6
         Left            =   645
         List            =   "FrmGuna.frx":0BF8
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3015
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1095
         TabIndex        =   0
         Text            =   "User"
         Top             =   615
         Width           =   2160
      End
      Begin AIFCmp1.asxToolbar Asx1 
         Height          =   420
         Left            =   3300
         Top             =   555
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   741
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
         BackColor       =   12632256
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
         Enabled         =   0   'False
         ButtonCount     =   1
         BackStyle       =   0
         ButtonKey1      =   "adduser"
         ButtonPicture1  =   "FrmGuna.frx":0C1B
         ButtonToolTipText1=   "Tambah pengguna"
      End
      Begin VB.Label LblCrnc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   15
         Top             =   2310
         Width           =   495
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type :"
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
         Index           =   1
         Left            =   285
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   2895
         TabIndex        =   11
         Top             =   3090
         Width           =   585
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   1530
         TabIndex        =   10
         Top             =   3060
         Width           =   390
      End
      Begin VB.Label Lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
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
         Index           =   0
         Left            =   285
         TabIndex        =   9
         Top             =   645
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmGuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sambung As Boolean

Private Sub Asx1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    'FrmPlview.Left = FrmGuna.Left + Asx1.Left
    'FrmPlview.Top = FrmGuna.Top + Asx1.Top - FrmPlview.Height
    'FrmPlview.Show
End Sub


Private Sub Form_Load()
    NamaPc.Caption = SelText
    LblCrnc = Crnc
    Combo3.List(0) = VS(2)
    Combo3.Text = VS(2)
    
    If Combo3.ListCount = 1 Then
        For s = 0 To uSDBe.DataCount("skema") - 1
            Combo3.AddItem uSDBe.DataGet("skema", "skema", s)
        Next s
    End If

    'untuk orang yang ingin sambung...
    If Sambung = True Then
        Text1 = SelAgn.CustomerName
        Combo3.Text = SelAgn.CustomerType
        Text1.Enabled = False
        Asx1.Enabled = False
    End If
End Sub



Private Sub MainBtn_Click(Index As Integer)

End Sub

Private Sub MainBtnKo_Click()
    FrmGuna.Hide
    Unload FrmGuna
End Sub

Private Sub MainBtnOk_Click()
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Flag Dalam Tag
    '
    ' Flag digunakan bagi mengenal pasti jenis dan keadaan
    ' pelanggan, samada prepaid,payg dan fixed time.. dan
    ' juga untuk mengenalpasti sama keadaan sambung adalah benar..
    ' semua flag dan nilai.. akan diletakkan dalam subitems(1)
    ' Bagi prepaid, nilai wang yang telah dibayar akan masukkan ke
    ' dalam tag bersama flag, bagi fixed time pula.. jumlah masa yang
    ' telah ditetapkan akan dimaukkan bersama flag juga, dan untuk payg hanya
    ' flag g sahaja..
    '
    '   Contoh :
    '       f10 = Fixed Time dan 10 minit
    '       p0.5 = Prepaid dan 0.5 sen
    '
    '   g - Pay As You Go
    '   P - PrePaid
    '   f - Fixed Time
    '
    
    '------------- variable declaration -------------
    Dim Cb1 As ComboBox, Cb2 As ComboBox
    Dim CsmerType As String, gMin As Integer, gJam As Integer

    '------------- defining -------------
    Set Cb1 = Combo1
    Set Cb2 = Combo2
    
    '------------- assigning & checking value -------------
    CsmerType = Combo3                                                                      'ambil jenis pelanggan
    If Text1.Text = "" Then Exit Sub                                                        'jika nama pelanggan kosong.. keluar..
    If Opt1(2).Value = True And IsNumeric(cbPaid) = False Then Exit Sub       'jika nilai dibayar bukan nombor.. keluar
    If Opt1(3).Value = True And (Cb1 = "" And Cb2 = "") Then Exit Sub          'jika nilai minit atau tidak dipilih.. keluar
    If CsmerType = "" Then CsmerType = VS(2)                                         'jika tiada jenis pelanggan dipilih.. guna jenis biasa
    
    
    ' an option ! yes its an option .. -------------
    '------- PAY AS YOU GO --------------------------------------
    If Opt1(1).Value = True Then
        SelAgn.CusStartPAYG UCase(Text1), CsmerType
        GoTo UnloadAll
    
    '------- PREPAID --------------------------------------------
    ElseIf Opt1(2).Value = True Then
        SelAgn.CusStartPPAID UCase(Text1), CsmerType, cbPaid, Sambung
        GoTo UnloadAll
        
    '------ FIXED TIME -----------------------------------------
    ElseIf Opt1(3).Value = True Then
        gJam = IIf(Cb1 = "", 0, Cb1)    'jika jam = "" return 0
        gMin = IIf(Cb2 = "", 0, Cb2)    'jika minit = "" return 0
        SelAgn.CusStartTIME UCase(Text1), CsmerType, gJam, gMin
        GoTo UnloadAll
    End If
Exit Sub
'---------------- End point of algorithm... -------------------

UnloadAll:

    FrmGuna.Hide
    If Sambung = True Then Sambung = False
    If SelTag <> "dump" Then SelAgn.NetSend "//kunci:0"
    Call UpdatePanel(SelText)
    Unload FrmGuna
End Sub

Private Sub Opt1_Click(Index As Integer)
    Select Case Index
    Case 1
        cbPaid.Enabled = False
        Combo1.Enabled = False
        Combo2.Enabled = False
    Case 2
        cbPaid.Enabled = True
        Combo1.Enabled = False
        Combo2.Enabled = False
    Case 3
        cbPaid.Enabled = False
        Combo1.Enabled = True
        Combo2.Enabled = True
    End Select
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveFrm Me.hwnd
End Sub

Public Sub FastLogin()
    Call AsxOk_ButtonClick(1, "ok")
End Sub
