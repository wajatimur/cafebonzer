VERSION 5.00
Begin VB.Form FrmInfo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CafeBonzer Info"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      DrawWidth       =   2
      FillColor       =   &H00808080&
      Height          =   945
      Left            =   45
      ScaleHeight     =   885
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   1950
      Width           =   4395
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright : Nematix Technology"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   555
         TabIndex        =   2
         Top             =   450
         Width           =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Author : Azri Jamil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1005
         TabIndex        =   1
         Top             =   150
         Width           =   2250
      End
   End
End
Attribute VB_Name = "FrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    FrmInfo.Hide
    Set FrmInfo = Nothing
    Unload FrmInfo
    FrmMain.Enabled = True
    FrmMain.SetFocus
End Sub

Private Sub Form_Load()
    lblnama = AmbilSet("namadaftar")
    lblkedai = AmbilSet("namacc")
    lblemail = AmbilSet("emailpengguna")
    If AmbilSet("demo") = "True" Then FrmMain.Caption = FrmMain.Caption & " UNREGISTERED"
End Sub
