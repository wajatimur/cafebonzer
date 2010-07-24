VERSION 5.00
Begin VB.Form FrmSuspen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CafeBonzerAG.Label3D Label3D1 
      Height          =   510
      Left            =   1800
      TabIndex        =   1
      Top             =   780
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   900
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777152
      ForeColor2      =   12632064
      Caption         =   "Y/N"
      BackColor       =   4210752
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Masa anda telah tamat ! Anda masih ingin menggunakan komputer?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   570
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   4305
   End
End
Attribute VB_Name = "FrmSuspen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyY Then
        FrmHost.Sock.SendData "nokguna"
        FrmSuspen.Hide: Kunci 0
        Set FrmSuspen = Nothing
    End If
    If KeyCode = vbKeyN Then
        FrmHost.Sock.SendData "takmboh"
        FrmSuspen.Hide: Set FrmSuspen = Nothing
    End If
End Sub

Private Sub Form_Load()
    PutOnTop Me.hwnd
End Sub
