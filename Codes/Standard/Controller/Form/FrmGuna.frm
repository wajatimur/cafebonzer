VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3750
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
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   3750
      TabIndex        =   12
      Top             =   0
      Width           =   3750
      Begin CafeBonzer.Label3D NamaPc 
         Height          =   300
         Left            =   1035
         TabIndex        =   13
         ToolTipText     =   "Nama Pc"
         Top             =   330
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Left            =   3225
         TabIndex        =   14
         Top             =   30
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "FrmGuna.frx":0000
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   15
      TabIndex        =   5
      Top             =   600
      Width           =   3720
      Begin CafeBonzer.Line3D uLine3D1 
         Height          =   45
         Left            =   60
         TabIndex        =   9
         Top             =   1170
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   79
         horizon         =   -1  'True
      End
      Begin VB.ComboBox cbPaid 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmGuna.frx":081B
         Left            =   2430
         List            =   "FrmGuna.frx":082E
         TabIndex        =   4
         Text            =   "1.00"
         Top             =   1695
         Width           =   1080
      End
      Begin VB.OptionButton Opt1 
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
         Left            =   270
         TabIndex        =   3
         Top             =   1740
         Width           =   1140
      End
      Begin VB.OptionButton Opt1 
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
         Left            =   270
         TabIndex        =   2
         Top             =   1305
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "FrmGuna.frx":0850
         Left            =   1905
         List            =   "FrmGuna.frx":0852
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox Text1 
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
         Top             =   300
         Width           =   2400
      End
      Begin VB.Label LblCrnc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RM"
         Height          =   195
         Left            =   2010
         TabIndex        =   8
         Top             =   1755
         Width           =   270
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
         TabIndex        =   7
         Top             =   765
         Width           =   1455
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
         TabIndex        =   6
         Top             =   330
         Width           =   630
      End
   End
   Begin CafeBonzer.XpButton BtnMenu 
      Height          =   450
      Index           =   0
      Left            =   2670
      TabIndex        =   10
      ToolTipText     =   "Add new employee."
      Top             =   2880
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   794
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
      MICON           =   "FrmGuna.frx":0854
      PICN            =   "FrmGuna.frx":0870
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton BtnMenu 
      Height          =   450
      Index           =   1
      Left            =   1665
      TabIndex        =   11
      ToolTipText     =   "Add new employee."
      Top             =   2880
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   794
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
      MICON           =   "FrmGuna.frx":0E0A
      PICN            =   "FrmGuna.frx":0E26
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
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmLogin
'    Project    : CafeBonzer
'
'    Description: Customer Login
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public CustomerContinue As Boolean
Private CustomerId As String


Private Sub BtnMenu_Click(Index As Integer)
    '------------- variable declaration -------------
    Dim CsmerType As String

    '------------- defining -------------
    If Index = 1 Then GoTo UnloadMe
    
    '------------- assigning & checking value -------------
    CsmerType = Combo3
    If Text1.Text = "" Then Exit Sub
    If Opt1(2).Value = True And IsNumeric(cbPaid) = False Then Exit Sub
    
    '------- PAY AS YOU GO --------------------------------------
    If Opt1(1).Value = True Then
        AgentSel.CusStartPAYG CustomerId, UCase$(Text1), CsmerType
        GoTo UnloadAll
    
    '------- PREPAID --------------------------------------------
    ElseIf Opt1(2).Value = True Then
        AgentSel.CusStartPPAID CustomerId, UCase$(Text1), CsmerType, cbPaid, CustomerContinue
        GoTo UnloadAll
        
    End If
Exit Sub
'---------------- End point of algorithm... -------------------

UnloadAll:
    If CustomerContinue = True Then CustomerContinue = False
    AgentSel.Commands.ConLogin (LogIn)
    AgentSel.AgnRecover
    Call UpdatePanel(SelText)

UnloadMe:
    Unload Me
End Sub


Private Sub Form_Load()
    CustomerId = "CU000000"
    NamaPc.Caption = SelText
    LblCrnc = Crnc
    
    For s = 0 To CDataSe.DataCount("PriceScheme") - 1
        Combo3.AddItem CDataSe.DataGet("PriceScheme", "Scheme", s)
    Next
    If Combo3.ListCount > 0 Then
        Combo3.ListIndex = 0
    End If
    
    If CustomerContinue = True Then
        Text1 = AgentSel.CustomerName
        Combo3.Text = AgentSel.CustomerType
        Text1.Enabled = False
        BtnKo.Enabled = False
    End If
End Sub


Private Sub Opt1_Click(Index As Integer)
    Select Case Index
    Case 1
        cbPaid.Enabled = False
    Case 2
        cbPaid.Enabled = True
    End Select
End Sub


Public Sub FastLogin()
    Call BtnMenu_Click(0)
End Sub
