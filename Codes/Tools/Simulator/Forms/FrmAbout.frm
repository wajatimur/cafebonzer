VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3420
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
   Icon            =   "FrmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed By :  Azri Jamil"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   660
      TabIndex        =   2
      Top             =   1110
      Width           =   2130
   End
   Begin VB.Label lblBuild 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   1185
      TabIndex        =   1
      Top             =   855
      Width           =   945
   End
   Begin VB.Image imgBg 
      Height          =   900
      Index           =   1
      Left            =   -30
      Picture         =   "FrmAbout.frx":000C
      Top             =   885
      Width           =   4650
   End
   Begin VB.Label lblCopy 
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
      Left            =   375
      TabIndex        =   0
      Top             =   555
      Width           =   2985
   End
   Begin VB.Image imgLogo 
      Height          =   420
      Left            =   195
      Picture         =   "FrmAbout.frx":0163
      Top             =   120
      Width           =   3000
   End
   Begin VB.Image imgBg 
      Height          =   900
      Index           =   0
      Left            =   -15
      Picture         =   "FrmAbout.frx":13BE
      Top             =   -15
      Width           =   4650
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lblBuild = "Build " & App.Major & "." & s_appBuild
End Sub

Private Sub imgBg_Click(Index As Integer)
    Unload FrmAbout
End Sub
