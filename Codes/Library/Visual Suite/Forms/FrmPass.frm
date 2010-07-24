VERSION 5.00
Begin VB.Form FrmPass 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1935
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4050
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
   Icon            =   "FrmPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VisualSuite.uLine3D uLine3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   735
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4050
      TabIndex        =   6
      Top             =   0
      Width           =   4050
      Begin VB.Image Image1 
         Height          =   960
         Left            =   3135
         Picture         =   "FrmPass.frx":000C
         Top             =   -15
         Width           =   960
      End
      Begin VB.Label LblCopy 
         BackStyle       =   0  'Transparent
         Caption         =   "Nematix Technology© 1996-2004"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   510
         Width           =   2730
      End
      Begin VB.Image imgLogo 
         Height          =   450
         Left            =   120
         Picture         =   "FrmPass.frx":0A0A
         Top             =   90
         Width           =   2700
      End
   End
   Begin VB.PictureBox Line3D2 
      Height          =   45
      Left            =   15
      ScaleHeight     =   45
      ScaleWidth      =   4020
      TabIndex        =   5
      Top             =   2070
      Width           =   4020
   End
   Begin VB.PictureBox Line3D1 
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   4050
      TabIndex        =   4
      Top             =   855
      Width           =   4050
   End
   Begin VB.TextBox TxtPass 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1425
      PasswordChar    =   "l"
      TabIndex        =   1
      ToolTipText     =   "Press ESC to exit.."
      Top             =   1395
      Width           =   2400
   End
   Begin VB.TextBox TxtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1425
      TabIndex        =   0
      Top             =   915
      Width           =   2400
   End
   Begin VB.Label LblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label LblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If BlnVsPasswordOnly = True Then
        TxtUser = StrVsCheckPassUser
        TxtUser.Enabled = False
    End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl = TxtUser Then
            TxtPass.SetFocus
        Else
            StrVsCheckPassUser = TxtUser
            StrVsCheckPassPassword = TxtPass
            Unload Me
        End If
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
