VERSION 5.00
Begin VB.Form FrmAppPass 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2490
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
   ScaleHeight     =   2490
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4050
      TabIndex        =   7
      Top             =   0
      Width           =   4050
      Begin VB.Image Image1 
         Height          =   960
         Left            =   3090
         Picture         =   "FrmPass.frx":000C
         Top             =   30
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
         TabIndex        =   8
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
   Begin CafeBonzer.Line3D Line3D2 
      Height          =   45
      Left            =   15
      TabIndex        =   6
      Top             =   2070
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CafeBonzer.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   5
      Top             =   855
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CafeBonzer.Label3D LblInfo 
      Height          =   210
      Left            =   1830
      TabIndex        =   2
      Top             =   2220
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   16777215
      ForeColor2      =   4210752
      Caption         =   "Please provide login info."
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
      Top             =   1545
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
      Top             =   1065
      Width           =   2400
   End
   Begin VB.Label LblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1590
      Width           =   1050
   End
   Begin VB.Label LblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Top             =   1110
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAppPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmAppPass
'    Project    : CafeBonzer
'
'    Description: Login Prompt
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If SecPasswordCheck(TxtUser, TxtPass) = True Then
            Unload FrmAppPass
            FrmMain.Show
            Exit Sub
        Else
            LblInfo.Caption = ST(1, 4)
            TxtPass = ""
            Exit Sub
        End If
    End If
    If KeyCode = vbKeyEscape Then
        AppExit False
    End If
End Sub
