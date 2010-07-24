VERSION 5.00
Begin VB.Form FrmPass 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2355
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CafeBonzer.Label3D LblInfo 
      Height          =   210
      Left            =   75
      TabIndex        =   4
      Top             =   2070
      Width           =   2805
      _ExtentX        =   4948
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
      BackColor       =   12632256
   End
   Begin CafeBonzer.Line3D Line3D 
      Height          =   45
      Left            =   30
      TabIndex        =   3
      Top             =   1950
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   79
      horizon         =   -1  'True
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
      Left            =   1470
      PasswordChar    =   "l"
      TabIndex        =   1
      ToolTipText     =   "Press ESC to exit.."
      Top             =   1470
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
      Left            =   1470
      TabIndex        =   0
      Top             =   990
      Width           =   2400
   End
   Begin VB.Label LblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   285
      Left            =   405
      TabIndex        =   7
      Top             =   1515
      Width           =   1050
   End
   Begin VB.Label LblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      Height          =   285
      Left            =   345
      TabIndex        =   6
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Label LblBuild 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Build 1.7.42"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2910
      TabIndex        =   5
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Image imgLogo 
      Height          =   420
      Left            =   585
      Picture         =   "FrmPass.frx":000C
      Top             =   90
      Width           =   3000
   End
   Begin VB.Label LblCopy 
      BackStyle       =   0  'Transparent
      Caption         =   "Nematix Technology© 1996-2002"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   555
      Width           =   2730
   End
   Begin VB.Image imgPass 
      Height          =   480
      Left            =   75
      Picture         =   "FrmPass.frx":1267
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgBg 
      Height          =   900
      Left            =   -15
      Picture         =   "FrmPass.frx":32D9
      Top             =   -45
      Width           =   4650
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If CekPass(TxtUser, TxtPass) = True Then
            Unload FrmPass
            FrmMain.Show
            Exit Sub
        Else
            LblInfo.Caption = "Access denied !"
            TxtPass = ""
            Exit Sub
        End If
    End If
    If KeyCode = vbKeyEscape Then
        Keluar False
    End If
End Sub

Private Sub Form_Load()
    LblBuild = "Build " & App.Major & "." & CbAppBuild
End Sub
